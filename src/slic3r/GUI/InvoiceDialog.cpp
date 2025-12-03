#include "InvoiceDialog.hpp"

#include <wx/notebook.h>
#include <wx/statline.h>
#include <wx/filedlg.h>
#include <wx/textfile.h>
#include <wx/msgdlg.h>
#include <wx/wfstream.h>
#include <wx/txtstrm.h>
#include <sstream>
#include <iomanip>
#include <regex>
#include <numeric>

#include "I18N.hpp"
#include "GUI_App.hpp"
#include "libslic3r/Print.hpp"
#include "libslic3r/AppConfig.hpp"
#include "libslic3r/PresetBundle.hpp"

namespace Slic3r {
namespace GUI {

InvoiceDialog::InvoiceDialog(wxWindow* parent, const PrintStatistics* stats)
    : DPIDialog(parent, wxID_ANY, _L("Invoice Generator - 3D Print Cost Calculator"),
                wxDefaultPosition, wxSize(900, 800),
                wxDEFAULT_DIALOG_STYLE | wxRESIZE_BORDER)
    , m_stats(stats)
{
    SetFont(wxGetApp().normal_font());
    
#ifdef __WINDOWS__
    wxGetApp().UpdateDarkUI(this);
#endif

    // Initialize default job profile
    m_current_job.parts_per_plate = 1;
    m_current_job.num_plates = 1;
    m_current_job.failure_rate = 5.0;
    m_current_job.labor_rate = 20.0;
    m_current_job.prep_time = 15.0;
    m_current_job.setup_time = 10.0;
    m_current_job.finishing_per_part = 5.0;
    m_current_job.finishing_per_plate = 0.0;
    m_current_job.printer_cost = 300.0;
    m_current_job.printer_lifespan = 15000.0;
    m_current_job.maintenance_cost = 0.10;
    m_current_job.power_watts = 130.0;
    m_current_job.electricity_cost = 0.15;
    m_current_job.bed_cost = 30.0;
    m_current_job.bed_lifespan = 5000.0;
    m_current_job.nozzle_cost = 2.0;
    m_current_job.nozzle_lifespan_kg = 25.0;
    m_current_job.solvent_cost = 0.0;
    m_current_job.solving_time = 0.0;
    m_current_job.tank_power = 0.0;
    m_current_job.finishing_materials = 0.0;
    m_current_job.markup_percent = 50.0;

    populate_filament_data();
    build_ui();
    load_global_settings();
    
    calculate_costs();

    Centre(wxBOTH);
}

void InvoiceDialog::on_dpi_changed(const wxRect& suggested_rect)
{
    const int em = em_unit();
    msw_buttons_rescale(this, em, { wxID_OK, wxID_CANCEL });
    Fit();
    Refresh();
}

wxString InvoiceDialog::format_time(const std::string& time_str) const
{
    if (time_str.empty())
        return _L("N/A");
    return wxString::FromUTF8(time_str);
}

double InvoiceDialog::parse_time_to_hours(const std::string& time_str) const
{
    double hours = 0.0;
    std::regex day_regex("(\\d+)\\s*d");
    std::regex hour_regex("(\\d+)\\s*h");
    std::regex min_regex("(\\d+)\\s*m");
    std::regex sec_regex("(\\d+)\\s*s");
    std::smatch match;
    
    if (std::regex_search(time_str, match, day_regex)) hours += std::stod(match[1]) * 24.0;
    if (std::regex_search(time_str, match, hour_regex)) hours += std::stod(match[1]);
    if (std::regex_search(time_str, match, min_regex)) hours += std::stod(match[1]) / 60.0;
    if (std::regex_search(time_str, match, sec_regex)) hours += std::stod(match[1]) / 3600.0;
    
    return hours;
}

void InvoiceDialog::populate_filament_data()
{
    m_filament_data.clear();
    if (!m_stats) return;

    PresetBundle* presets = wxGetApp().preset_bundle;
    if (!presets) return;

    for (auto const& [extruder_id, usage] : m_stats->filament_stats) {
        FilamentData data;
        data.extruder_id = extruder_id;
        
        std::string filament_name = "Unknown";
        std::string color = "#808080";
        double cost_per_kg = 20.0;
        double density = 1.24;
        double diameter = 1.75;
        
        auto get_val = [&](const std::string& key, size_t idx) -> std::string {
            auto opt = presets->project_config.option(key);
            if (!opt) return "";

            if (auto opt_str = dynamic_cast<const ConfigOptionStrings*>(opt)) {
                if (idx < opt_str->values.size()) return opt_str->values[idx];
            }
            else if (auto opt_floats = dynamic_cast<const ConfigOptionFloats*>(opt)) {
                if (idx < opt_floats->values.size()) return std::to_string(opt_floats->values[idx]);
            }
            else if (auto opt_ints = dynamic_cast<const ConfigOptionInts*>(opt)) {
                if (idx < opt_ints->values.size()) return std::to_string(opt_ints->values[idx]);
            }
            return "";
        };

        if (extruder_id < presets->filament_presets.size()) {
            filament_name = presets->filament_presets[extruder_id];
        } else {
            filament_name = "Filament " + std::to_string(extruder_id + 1);
        }
        
        std::string col = get_val("filament_colour", extruder_id);
        if (!col.empty()) color = col;
        
        std::string cost_str = get_val("filament_cost", extruder_id);
        if (!cost_str.empty()) cost_per_kg = std::stod(cost_str);
        
        std::string dens_str = get_val("filament_density", extruder_id);
        if (!dens_str.empty()) density = std::stod(dens_str);
        
        std::string diam_str = get_val("filament_diameter", extruder_id);
        if (!diam_str.empty()) diameter = std::stod(diam_str);
        
        data.name = filament_name;
        data.color = color;
        data.cost_per_kg = cost_per_kg;
        
        double radius = diameter / 2.0;
        double area = 3.14159265359 * radius * radius;
        double volume_mm3 = usage * area; 
        data.weight_g = volume_mm3 * density / 1000.0; 
        
        if (m_stats->filament_stats.size() == 1 && m_stats->total_weight > 0) {
             data.weight_g = m_stats->total_weight;
        }

        data.calculated_cost = (data.weight_g / 1000.0) * data.cost_per_kg;
        
        m_filament_data.push_back(data);
    }
    
    if (m_filament_data.empty() && m_stats->total_weight > 0) {
        FilamentData data;
        data.extruder_id = 0;
        data.name = "Default Filament";
        data.color = "#808080";
        data.weight_g = m_stats->total_weight;
        data.cost_per_kg = 20.0;
        data.calculated_cost = (data.weight_g / 1000.0) * data.cost_per_kg;
        m_filament_data.push_back(data);
    }
}

void InvoiceDialog::build_ui()
{
    wxBoxSizer* main_sizer = new wxBoxSizer(wxVERTICAL);
    
    wxNotebook* notebook = new wxNotebook(this, wxID_ANY);
    
    build_customer_info_tab(notebook);
    build_job_info_tab(notebook);
    build_materials_tab(notebook);
    build_labor_tab(notebook);
    build_machine_tab(notebook);
    build_tooling_tab(notebook);
    build_postprocess_tab(notebook);
    build_markup_tab(notebook);
    build_results_tab(notebook);
    
    main_sizer->Add(notebook, 1, wxEXPAND | wxALL, 5);
    
    wxBoxSizer* btn_sizer = new wxBoxSizer(wxHORIZONTAL);
    
    m_btn_calculate = new wxButton(this, wxID_ANY, _L("Calculate"));
    m_btn_save_job = new wxButton(this, wxID_ANY, _L("Save Job Profile"));
    m_btn_export = new wxButton(this, wxID_ANY, _L("Export Invoice (Excel)"));
    m_btn_close = new wxButton(this, wxID_CANCEL, _L("Close"));
    
    btn_sizer->Add(m_btn_calculate, 0, wxALL, 5);
    btn_sizer->Add(m_btn_save_job, 0, wxALL, 5);
    btn_sizer->Add(m_btn_export, 0, wxALL, 5);
    btn_sizer->AddStretchSpacer();
    btn_sizer->Add(m_btn_close, 0, wxALL, 5);
    
    main_sizer->Add(btn_sizer, 0, wxEXPAND | wxALL, 5);
    
    SetSizer(main_sizer);
    
    m_btn_calculate->Bind(wxEVT_BUTTON, &InvoiceDialog::on_calculate, this);
    m_btn_save_job->Bind(wxEVT_BUTTON, &InvoiceDialog::on_save_job, this);
    m_btn_export->Bind(wxEVT_BUTTON, &InvoiceDialog::on_export_invoice, this);
    
    Fit();
}

void InvoiceDialog::build_customer_info_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    auto add_field = [&](const wxString& label, wxTextCtrl*& ctrl) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxTextCtrl(panel, wxID_ANY);
        grid->Add(ctrl, 1, wxEXPAND);
    };
    
    add_field(_L("My Business Name:"), m_txt_business_name);
    add_field(_L("Customer Name:"), m_txt_customer_name);
    add_field(_L("Customer Email:"), m_txt_customer_email);
    add_field(_L("Customer Phone:"), m_txt_customer_phone);
    add_field(_L("Job Name:"), m_txt_job_name);
    add_field(_L("Job Description:"), m_txt_job_description);
    
    grid->Add(new wxStaticText(panel, wxID_ANY, _L("Saved Job Profiles:")), 0, wxALIGN_CENTER_VERTICAL);
    wxBoxSizer* profile_sizer = new wxBoxSizer(wxHORIZONTAL);
    m_combo_job_profiles = new wxComboBox(panel, wxID_ANY, "", wxDefaultPosition, wxDefaultSize, 0, NULL, wxCB_READONLY);
    m_btn_load_job = new wxButton(panel, wxID_ANY, _L("Load"));
    m_btn_delete_job = new wxButton(panel, wxID_ANY, _L("Delete"));
    
    profile_sizer->Add(m_combo_job_profiles, 1, wxEXPAND | wxRIGHT, 5);
    profile_sizer->Add(m_btn_load_job, 0, wxRIGHT, 5);
    profile_sizer->Add(m_btn_delete_job, 0);
    
    grid->Add(profile_sizer, 1, wxEXPAND);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Customer & Job"));
    
    m_btn_load_job->Bind(wxEVT_BUTTON, &InvoiceDialog::on_load_job, this);
    m_btn_delete_job->Bind(wxEVT_BUTTON, &InvoiceDialog::on_delete_job, this);
    
    refresh_job_profiles_combo();
}

void InvoiceDialog::build_job_info_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    grid->Add(new wxStaticText(panel, wxID_ANY, _L("Parts per Plate:")), 0, wxALIGN_CENTER_VERTICAL);
    m_parts_per_plate = new wxSpinCtrl(panel, wxID_ANY, "1", wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 1, 1000, 1);
    grid->Add(m_parts_per_plate, 1, wxEXPAND);
    
    grid->Add(new wxStaticText(panel, wxID_ANY, _L("Number of Plates:")), 0, wxALIGN_CENTER_VERTICAL);
    m_num_plates = new wxSpinCtrl(panel, wxID_ANY, "1", wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 1, 1000, 1);
    grid->Add(m_num_plates, 1, wxEXPAND);
    
    grid->Add(new wxStaticText(panel, wxID_ANY, _L("Failure Rate (%):")), 0, wxALIGN_CENTER_VERTICAL);
    m_failure_rate = new wxSpinCtrlDouble(panel, wxID_ANY, "5.0", wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, 50.0, 5.0, 1.0);
    grid->Add(m_failure_rate, 1, wxEXPAND);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    
    wxStaticBoxSizer* stats_box = new wxStaticBoxSizer(wxVERTICAL, panel, _L("Slicer Statistics"));
    wxFlexGridSizer* stats_grid = new wxFlexGridSizer(2, 10, 10);
    stats_grid->AddGrowableCol(1);
    
    stats_grid->Add(new wxStaticText(stats_box->GetStaticBox(), wxID_ANY, _L("Print Time:")), 0, wxALIGN_CENTER_VERTICAL);
    m_lbl_print_time = new wxStaticText(stats_box->GetStaticBox(), wxID_ANY, _L("N/A"));
    stats_grid->Add(m_lbl_print_time, 1, wxEXPAND);
    
    stats_grid->Add(new wxStaticText(stats_box->GetStaticBox(), wxID_ANY, _L("Total Weight:")), 0, wxALIGN_CENTER_VERTICAL);
    m_lbl_total_weight = new wxStaticText(stats_box->GetStaticBox(), wxID_ANY, _L("N/A"));
    stats_grid->Add(m_lbl_total_weight, 1, wxEXPAND);
    
    stats_box->Add(stats_grid, 1, wxEXPAND | wxALL, 10);
    sizer->Add(stats_box, 0, wxEXPAND | wxALL, 10);
    
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Job Parameters"));
}

void InvoiceDialog::build_materials_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    
    m_materials_grid = new wxGrid(panel, wxID_ANY);
    m_materials_grid->CreateGrid(0, 5);
    m_materials_grid->SetColLabelValue(0, _L("Filament"));
    m_materials_grid->SetColLabelValue(1, _L("Color"));
    m_materials_grid->SetColLabelValue(2, _L("Weight (g)"));
    m_materials_grid->SetColLabelValue(3, _L("Cost ($/kg)"));
    m_materials_grid->SetColLabelValue(4, _L("Total Cost"));
    
    m_materials_grid->SetColFormatFloat(2, 2, 2);
    m_materials_grid->SetColFormatFloat(3, 2, 2);
    m_materials_grid->SetColFormatFloat(4, 2, 2);
    
    m_materials_grid->AutoSizeColumns();
    
    sizer->Add(m_materials_grid, 1, wxEXPAND | wxALL, 10);
    
    wxBoxSizer* total_sizer = new wxBoxSizer(wxHORIZONTAL);
    total_sizer->Add(new wxStaticText(panel, wxID_ANY, _L("Total Material Cost: ")), 0, wxALIGN_CENTER_VERTICAL);
    m_lbl_total_material_cost = new wxStaticText(panel, wxID_ANY, "$0.00");
    m_lbl_total_material_cost->SetFont(m_lbl_total_material_cost->GetFont().Bold());
    total_sizer->Add(m_lbl_total_material_cost, 0, wxALIGN_CENTER_VERTICAL);
    
    sizer->Add(total_sizer, 0, wxALIGN_RIGHT | wxALL, 10);
    
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Materials"));
    
    update_materials_grid();
    
    m_materials_grid->Bind(wxEVT_GRID_CELL_CHANGED, &InvoiceDialog::on_filament_cost_changed, this);
}

void InvoiceDialog::build_labor_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    auto add_spin = [&](const wxString& label, wxSpinCtrlDouble*& ctrl, double val, double max) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxSpinCtrlDouble(panel, wxID_ANY, wxString::Format("%.2f", val), wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, max, val, 1.0);
        grid->Add(ctrl, 1, wxEXPAND);
    };
    
    add_spin(_L("Hourly Rate ($/hr):"), m_labor_rate, 20.0, 500.0);
    add_spin(_L("Slicing/Prep Time (min/plate):"), m_prep_time, 15.0, 120.0);
    add_spin(_L("Machine Setup (min/plate):"), m_setup_time, 10.0, 120.0);
    add_spin(_L("Finishing Time (min/part):"), m_finishing_per_part, 5.0, 120.0);
    add_spin(_L("Finishing Time (min/plate):"), m_finishing_per_plate, 0.0, 120.0);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Labor"));
}

void InvoiceDialog::build_machine_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    auto add_spin = [&](const wxString& label, wxSpinCtrlDouble*& ctrl, double val, double max, double step) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxSpinCtrlDouble(panel, wxID_ANY, wxString::Format("%.2f", val), wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, max, val, step);
        grid->Add(ctrl, 1, wxEXPAND);
    };
    
    add_spin(_L("Printer Cost ($):"), m_printer_cost, 300.0, 50000.0, 50.0);
    add_spin(_L("Printer Lifespan (hours):"), m_printer_lifespan, 15000.0, 100000.0, 1000.0);
    add_spin(_L("Maintenance Cost ($/hr):"), m_maintenance_cost, 0.10, 10.0, 0.01);
    add_spin(_L("Average Power (Watts):"), m_power_watts, 130.0, 2000.0, 10.0);
    add_spin(_L("Electricity Cost ($/kWh):"), m_electricity_cost, 0.15, 1.0, 0.01);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Machine"));
}

void InvoiceDialog::build_tooling_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    auto add_spin = [&](const wxString& label, wxSpinCtrlDouble*& ctrl, double val, double max, double step) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxSpinCtrlDouble(panel, wxID_ANY, wxString::Format("%.2f", val), wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, max, val, step);
        grid->Add(ctrl, 1, wxEXPAND);
    };
    
    add_spin(_L("Build Plate Cost ($):"), m_bed_cost, 30.0, 500.0, 5.0);
    add_spin(_L("Build Plate Lifespan (hours):"), m_bed_lifespan, 5000.0, 50000.0, 500.0);
    add_spin(_L("Nozzle Cost ($):"), m_nozzle_cost, 2.0, 200.0, 1.0);
    add_spin(_L("Nozzle Lifespan (kg):"), m_nozzle_lifespan_kg, 25.0, 500.0, 5.0);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Tooling"));
}

void InvoiceDialog::build_postprocess_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    auto add_spin = [&](const wxString& label, wxSpinCtrlDouble*& ctrl, double val, double max, double step) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxSpinCtrlDouble(panel, wxID_ANY, wxString::Format("%.2f", val), wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, max, val, step);
        grid->Add(ctrl, 1, wxEXPAND);
    };
    
    add_spin(_L("Solvent Cost ($/L):"), m_solvent_cost, 0.0, 100.0, 1.0);
    add_spin(_L("Solving Time (hours):"), m_solving_time, 0.0, 48.0, 0.5);
    add_spin(_L("Tank Power (Watts):"), m_tank_power, 0.0, 1000.0, 10.0);
    add_spin(_L("Finishing Materials ($/plate):"), m_finishing_materials, 0.0, 100.0, 1.0);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Post-Processing"));
}

void InvoiceDialog::build_markup_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 10);
    grid->AddGrowableCol(1);
    
    grid->Add(new wxStaticText(panel, wxID_ANY, _L("Markup (%):")), 0, wxALIGN_CENTER_VERTICAL);
    m_markup_percent = new wxSpinCtrlDouble(panel, wxID_ANY, "50.0", wxDefaultPosition, wxDefaultSize, wxSP_ARROW_KEYS, 0.0, 500.0, 50.0, 5.0);
    grid->Add(m_markup_percent, 1, wxEXPAND);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 10);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Markup"));
}

void InvoiceDialog::build_results_tab(wxNotebook* notebook)
{
    wxPanel* panel = new wxPanel(notebook);
    wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    
    wxFlexGridSizer* grid = new wxFlexGridSizer(2, 10, 5);
    grid->AddGrowableCol(1);
    
    auto add_row = [&](const wxString& label, wxStaticText*& ctrl, bool bold = false) {
        grid->Add(new wxStaticText(panel, wxID_ANY, label), 0, wxALIGN_CENTER_VERTICAL);
        ctrl = new wxStaticText(panel, wxID_ANY, "$0.00");
        if (bold) ctrl->SetFont(ctrl->GetFont().Bold());
        grid->Add(ctrl, 1, wxALIGN_RIGHT);
    };
    
    add_row(_L("Material Cost:"), m_lbl_material_cost);
    add_row(_L("Labor Cost:"), m_lbl_labor_cost);
    add_row(_L("Machine Cost:"), m_lbl_machine_cost);
    add_row(_L("Tooling Cost:"), m_lbl_tooling_cost);
    add_row(_L("Post-Processing Cost:"), m_lbl_postprocess_cost);
    
    grid->Add(new wxStaticLine(panel), 0, wxEXPAND | wxTOP | wxBOTTOM, 5);
    grid->Add(new wxStaticLine(panel), 0, wxEXPAND | wxTOP | wxBOTTOM, 5);
    
    add_row(_L("Subtotal (per plate):"), m_lbl_subtotal);
    add_row(_L("Failure Rate Adjustment:"), m_lbl_failure_adjustment);
    add_row(_L("Cost Per Part:"), m_lbl_cost_per_part);
    add_row(_L("Markup Amount:"), m_lbl_markup_amount);
    
    grid->Add(new wxStaticLine(panel), 0, wxEXPAND | wxTOP | wxBOTTOM, 5);
    grid->Add(new wxStaticLine(panel), 0, wxEXPAND | wxTOP | wxBOTTOM, 5);
    
    add_row(_L("FINAL PRICE PER PART:"), m_lbl_final_price, true);
    add_row(_L("TOTAL JOB COST:"), m_lbl_total_job_cost, true);
    
    sizer->Add(grid, 1, wxEXPAND | wxALL, 20);
    panel->SetSizer(sizer);
    notebook->AddPage(panel, _L("Results"));
}

void InvoiceDialog::update_materials_grid()
{
    if (m_materials_grid->GetNumberRows() > 0)
        m_materials_grid->DeleteRows(0, m_materials_grid->GetNumberRows());
        
    m_materials_grid->AppendRows(m_filament_data.size());
    
    for (size_t i = 0; i < m_filament_data.size(); ++i) {
        const auto& data = m_filament_data[i];
        m_materials_grid->SetCellValue(i, 0, data.name);
        m_materials_grid->SetCellValue(i, 1, data.color);
        m_materials_grid->SetCellValue(i, 2, wxString::Format("%.2f", data.weight_g));
        m_materials_grid->SetCellValue(i, 3, wxString::Format("%.2f", data.cost_per_kg));
        m_materials_grid->SetCellValue(i, 4, wxString::Format("%.2f", data.calculated_cost));
        
        m_materials_grid->SetReadOnly(i, 0);
        m_materials_grid->SetReadOnly(i, 1);
        m_materials_grid->SetReadOnly(i, 2);
        m_materials_grid->SetReadOnly(i, 4);
    }
    
    m_materials_grid->AutoSizeColumns();
}

void InvoiceDialog::on_filament_cost_changed(wxGridEvent& event)
{
    int row = event.GetRow();
    int col = event.GetCol();
    
    if (col == 3 && row >= 0 && row < m_filament_data.size()) {
        double new_cost = 0.0;
        wxString val = m_materials_grid->GetCellValue(row, col);
        if (val.ToDouble(&new_cost)) {
            m_filament_data[row].cost_per_kg = new_cost;
            m_filament_data[row].calculated_cost = (m_filament_data[row].weight_g / 1000.0) * new_cost;
            m_materials_grid->SetCellValue(row, 4, wxString::Format("%.2f", m_filament_data[row].calculated_cost));
            calculate_costs();
        }
    }
}

void InvoiceDialog::calculate_costs()
{
    double total_material_cost = 0.0;
    for (const auto& data : m_filament_data) {
        total_material_cost += data.calculated_cost;
    }
    m_lbl_total_material_cost->SetLabel(wxString::Format("$%.2f", total_material_cost));
    
    double print_time_hours = 0.0;
    if (m_stats) {
        print_time_hours = parse_time_to_hours(m_stats->estimated_normal_print_time);
        m_lbl_print_time->SetLabel(format_time(m_stats->estimated_normal_print_time));
        m_lbl_total_weight->SetLabel(wxString::Format("%.2f g", m_stats->total_weight));
    }
    
    int parts_per_plate = m_parts_per_plate->GetValue();
    int num_plates = m_num_plates->GetValue();
    double failure_rate = m_failure_rate->GetValue() / 100.0;
    
    double labor_rate = m_labor_rate->GetValue();
    double prep_time = m_prep_time->GetValue();
    double setup_time = m_setup_time->GetValue();
    double finishing_per_part = m_finishing_per_part->GetValue();
    double finishing_per_plate = m_finishing_per_plate->GetValue();
    
    double printer_cost = m_printer_cost->GetValue();
    double printer_lifespan = m_printer_lifespan->GetValue();
    double maintenance_cost = m_maintenance_cost->GetValue();
    double power_watts = m_power_watts->GetValue();
    double electricity_cost = m_electricity_cost->GetValue();
    
    double bed_cost = m_bed_cost->GetValue();
    double bed_lifespan = m_bed_lifespan->GetValue();
    double nozzle_cost = m_nozzle_cost->GetValue();
    double nozzle_lifespan_kg = m_nozzle_lifespan_kg->GetValue();
    
    double solvent_cost = m_solvent_cost->GetValue();
    double solving_time = m_solving_time->GetValue();
    double tank_power = m_tank_power->GetValue();
    double finishing_materials = m_finishing_materials->GetValue();
    
    double markup_percent = m_markup_percent->GetValue() / 100.0;
    
    m_calc_material_cost = total_material_cost;
    
    double labor_time_minutes = prep_time + setup_time + (finishing_per_part * parts_per_plate) + finishing_per_plate;
    m_calc_labor_cost = (labor_time_minutes / 60.0) * labor_rate;
    
    double depreciation_per_hour = printer_cost / printer_lifespan;
    double power_cost_per_hour = (power_watts / 1000.0) * electricity_cost;
    double machine_hourly_rate = depreciation_per_hour + maintenance_cost + power_cost_per_hour;
    m_calc_machine_cost = machine_hourly_rate * print_time_hours;
    
    double c_bed = (bed_cost / bed_lifespan) * print_time_hours;
    double total_filament_kg = 0.0;
    for(const auto& d : m_filament_data) total_filament_kg += d.weight_g / 1000.0;
    double c_nozzle = (nozzle_cost / nozzle_lifespan_kg) * total_filament_kg;
    m_calc_tooling_cost = c_bed + c_nozzle;
    
    double c_tank_power = (tank_power / 1000.0) * electricity_cost * solving_time;
    m_calc_postprocess_cost = c_tank_power + finishing_materials; 
    
    m_calc_subtotal = m_calc_material_cost + m_calc_labor_cost + m_calc_machine_cost + m_calc_tooling_cost + m_calc_postprocess_cost;
    
    m_calc_failure_adjustment = 0.0;
    if (failure_rate < 1.0) {
        m_calc_failure_adjustment = (m_calc_subtotal / (1.0 - failure_rate)) - m_calc_subtotal;
    }
    
    double total_plate_cost = m_calc_subtotal + m_calc_failure_adjustment;
    m_calc_cost_per_part = (parts_per_plate > 0) ? total_plate_cost / parts_per_plate : total_plate_cost;
    
    m_calc_markup_amount = m_calc_cost_per_part * markup_percent;
    m_calc_final_price = m_calc_cost_per_part + m_calc_markup_amount;
    
    int total_parts = parts_per_plate * num_plates;
    m_calc_total_job_cost = m_calc_final_price * total_parts;
    m_calc_print_time_hours = print_time_hours;
    
    m_lbl_material_cost->SetLabel(wxString::Format("$%.2f", m_calc_material_cost));
    m_lbl_labor_cost->SetLabel(wxString::Format("$%.2f", m_calc_labor_cost));
    m_lbl_machine_cost->SetLabel(wxString::Format("$%.2f", m_calc_machine_cost));
    m_lbl_tooling_cost->SetLabel(wxString::Format("$%.2f", m_calc_tooling_cost));
    m_lbl_postprocess_cost->SetLabel(wxString::Format("$%.2f", m_calc_postprocess_cost));
    m_lbl_subtotal->SetLabel(wxString::Format("$%.2f", m_calc_subtotal));
    m_lbl_failure_adjustment->SetLabel(wxString::Format("+$%.2f", m_calc_failure_adjustment));
    m_lbl_cost_per_part->SetLabel(wxString::Format("$%.2f", m_calc_cost_per_part));
    m_lbl_markup_amount->SetLabel(wxString::Format("+$%.2f", m_calc_markup_amount));
    m_lbl_final_price->SetLabel(wxString::Format("$%.2f", m_calc_final_price));
    m_lbl_total_job_cost->SetLabel(wxString::Format("$%.2f (%d parts)", m_calc_total_job_cost, total_parts));
}

void InvoiceDialog::on_calculate(wxCommandEvent& event)
{
    calculate_costs();
}

void InvoiceDialog::load_global_settings()
{
    AppConfig* config = wxGetApp().app_config;
    if (!config) return;
    
    m_txt_business_name->SetValue(config->get("invoice", "business_name"));
    
    std::string last_profile = config->get("invoice", "last_profile");
    if (!last_profile.empty()) {
        load_job_profile(last_profile);
        m_combo_job_profiles->SetValue(last_profile);
    }
}

void InvoiceDialog::save_global_settings()
{
    AppConfig* config = wxGetApp().app_config;
    if (!config) return;
    
    config->set("invoice", "business_name", m_txt_business_name->GetValue().ToStdString());
    config->set("invoice", "last_profile", m_combo_job_profiles->GetValue().ToStdString());
    config->save();
}

void InvoiceDialog::load_job_profile(const std::string& profile_name)
{
    AppConfig* config = wxGetApp().app_config;
    if (!config || profile_name.empty()) return;
    
    std::string prefix = "invoice_job_" + profile_name + "_";
    
    auto get_str = [&](const std::string& key) { return config->get(prefix + key); };
    auto get_int = [&](const std::string& key, int def) { 
        std::string v = get_str(key); return v.empty() ? def : std::stoi(v); 
    };
    auto get_dbl = [&](const std::string& key, double def) { 
        std::string v = get_str(key); return v.empty() ? def : std::stod(v); 
    };
    
    m_txt_customer_name->SetValue(get_str("customer_name"));
    m_txt_customer_email->SetValue(get_str("customer_email"));
    m_txt_customer_phone->SetValue(get_str("customer_phone"));
    m_txt_job_name->SetValue(get_str("job_name"));
    m_txt_job_description->SetValue(get_str("job_description"));
    
    m_parts_per_plate->SetValue(get_int("parts_per_plate", 1));
    m_num_plates->SetValue(get_int("num_plates", 1));
    m_failure_rate->SetValue(get_dbl("failure_rate", 5.0));
    
    m_labor_rate->SetValue(get_dbl("labor_rate", 20.0));
    m_prep_time->SetValue(get_dbl("prep_time", 15.0));
    m_setup_time->SetValue(get_dbl("setup_time", 10.0));
    m_finishing_per_part->SetValue(get_dbl("finishing_per_part", 5.0));
    m_finishing_per_plate->SetValue(get_dbl("finishing_per_plate", 0.0));
    
    m_printer_cost->SetValue(get_dbl("printer_cost", 300.0));
    m_printer_lifespan->SetValue(get_dbl("printer_lifespan", 15000.0));
    m_maintenance_cost->SetValue(get_dbl("maintenance_cost", 0.10));
    m_power_watts->SetValue(get_dbl("power_watts", 130.0));
    m_electricity_cost->SetValue(get_dbl("electricity_cost", 0.15));
    
    m_bed_cost->SetValue(get_dbl("bed_cost", 30.0));
    m_bed_lifespan->SetValue(get_dbl("bed_lifespan", 5000.0));
    m_nozzle_cost->SetValue(get_dbl("nozzle_cost", 2.0));
    m_nozzle_lifespan_kg->SetValue(get_dbl("nozzle_lifespan_kg", 25.0));
    
    m_solvent_cost->SetValue(get_dbl("solvent_cost", 0.0));
    m_solving_time->SetValue(get_dbl("solving_time", 0.0));
    m_tank_power->SetValue(get_dbl("tank_power", 0.0));
    m_finishing_materials->SetValue(get_dbl("finishing_materials", 0.0));
    
    m_markup_percent->SetValue(get_dbl("markup_percent", 50.0));
    
    calculate_costs();
}

void InvoiceDialog::save_job_profile(const std::string& profile_name)
{
    AppConfig* config = wxGetApp().app_config;
    if (!config || profile_name.empty()) return;
    
    std::string prefix = "invoice_job_" + profile_name + "_";
    
    auto set_val = [&](const std::string& key, const std::string& val) { config->set(prefix + key, val); };
    
    set_val("customer_name", m_txt_customer_name->GetValue().ToStdString());
    set_val("customer_email", m_txt_customer_email->GetValue().ToStdString());
    set_val("customer_phone", m_txt_customer_phone->GetValue().ToStdString());
    set_val("job_name", m_txt_job_name->GetValue().ToStdString());
    set_val("job_description", m_txt_job_description->GetValue().ToStdString());
    
    set_val("parts_per_plate", std::to_string(m_parts_per_plate->GetValue()));
    set_val("num_plates", std::to_string(m_num_plates->GetValue()));
    set_val("failure_rate", std::to_string(m_failure_rate->GetValue()));
    
    set_val("labor_rate", std::to_string(m_labor_rate->GetValue()));
    set_val("prep_time", std::to_string(m_prep_time->GetValue()));
    set_val("setup_time", std::to_string(m_setup_time->GetValue()));
    set_val("finishing_per_part", std::to_string(m_finishing_per_part->GetValue()));
    set_val("finishing_per_plate", std::to_string(m_finishing_per_plate->GetValue()));
    
    set_val("printer_cost", std::to_string(m_printer_cost->GetValue()));
    set_val("printer_lifespan", std::to_string(m_printer_lifespan->GetValue()));
    set_val("maintenance_cost", std::to_string(m_maintenance_cost->GetValue()));
    set_val("power_watts", std::to_string(m_power_watts->GetValue()));
    set_val("electricity_cost", std::to_string(m_electricity_cost->GetValue()));
    
    set_val("bed_cost", std::to_string(m_bed_cost->GetValue()));
    set_val("bed_lifespan", std::to_string(m_bed_lifespan->GetValue()));
    set_val("nozzle_cost", std::to_string(m_nozzle_cost->GetValue()));
    set_val("nozzle_lifespan_kg", std::to_string(m_nozzle_lifespan_kg->GetValue()));
    
    set_val("solvent_cost", std::to_string(m_solvent_cost->GetValue()));
    set_val("solving_time", std::to_string(m_solving_time->GetValue()));
    set_val("tank_power", std::to_string(m_tank_power->GetValue()));
    set_val("finishing_materials", std::to_string(m_finishing_materials->GetValue()));
    
    set_val("markup_percent", std::to_string(m_markup_percent->GetValue()));
    
    std::string profiles_str = config->get("invoice_profiles");
    std::vector<std::string> profiles;
    std::stringstream ss(profiles_str);
    std::string item;
    bool found = false;
    while (std::getline(ss, item, ';')) {
        if (item == profile_name) found = true;
        profiles.push_back(item);
    }
    if (!found) {
        profiles.push_back(profile_name);
        std::string new_list;
        for (const auto& p : profiles) new_list += p + ";";
        config->set("invoice_profiles", new_list);
    }
    
    config->save();
    refresh_job_profiles_combo();
}

void InvoiceDialog::delete_job_profile(const std::string& profile_name)
{
    AppConfig* config = wxGetApp().app_config;
    if (!config || profile_name.empty()) return;
    
    std::string profiles_str = config->get("invoice_profiles");
    std::string new_list;
    std::stringstream ss(profiles_str);
    std::string item;
    while (std::getline(ss, item, ';')) {
        if (item != profile_name) new_list += item + ";";
    }
    config->set("invoice_profiles", new_list);
    config->save();
    refresh_job_profiles_combo();
}

void InvoiceDialog::refresh_job_profiles_combo()
{
    m_combo_job_profiles->Clear();
    AppConfig* config = wxGetApp().app_config;
    if (!config) return;
    
    std::string profiles_str = config->get("invoice_profiles");
    std::stringstream ss(profiles_str);
    std::string item;
    while (std::getline(ss, item, ';')) {
        if (!item.empty()) m_combo_job_profiles->Append(item);
    }
}

void InvoiceDialog::on_save_job(wxCommandEvent& event)
{
    wxTextEntryDialog dlg(this, _L("Enter name for this job profile:"), _L("Save Job Profile"));
    if (dlg.ShowModal() == wxID_OK) {
        save_job_profile(dlg.GetValue().ToStdString());
        save_global_settings();
    }
}

void InvoiceDialog::on_load_job(wxCommandEvent& event)
{
    wxString val = m_combo_job_profiles->GetValue();
    if (!val.IsEmpty()) {
        load_job_profile(val.ToStdString());
    }
}

void InvoiceDialog::on_delete_job(wxCommandEvent& event)
{
    wxString val = m_combo_job_profiles->GetValue();
    if (!val.IsEmpty()) {
        if (wxMessageBox(_L("Are you sure you want to delete this profile?"), _L("Confirm Delete"), wxYES_NO | wxICON_QUESTION) == wxYES) {
            delete_job_profile(val.ToStdString());
            m_combo_job_profiles->SetValue("");
        }
    }
}

std::string InvoiceDialog::escape_xml(const std::string& str) const {
    std::string result;
    for (char c : str) {
        switch (c) {
            case '&': result += "&amp;"; break;
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default: result += c;
        }
    }
    return result;
}

void InvoiceDialog::export_to_excel(const wxString& path)
{
    wxFFileOutputStream output(path);
    if (!output.IsOk()) return;
    wxTextOutputStream file(output);
    
    file << "<?xml version=\"1.0\"?>\n";
    file << "<?mso-application progid=\"Excel.Sheet\"?>\n";
    file << "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n";
    file << " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n";
    file << " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n";
    file << " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n";
    file << " xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n";
    
    file << "<Styles>\n";
    file << " <Style ss:ID=\"Default\" ss:Name=\"Normal\">\n";
    file << "  <Alignment ss:Vertical=\"Bottom\"/>\n";
    file << "  <Borders/>\n";
    file << "  <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>\n";
    file << "  <Interior/>\n";
    file << "  <NumberFormat/>\n";
    file << "  <Protection/>\n";
    file << " </Style>\n";
    file << " <Style ss:ID=\"sHeader\">\n";
    file << "  <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"14\" ss:Bold=\"1\"/>\n";
    file << "  <Alignment ss:Horizontal=\"Center\"/>\n";
    file << " </Style>\n";
    file << " <Style ss:ID=\"sBold\">\n";
    file << "  <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Bold=\"1\"/>\n";
    file << " </Style>\n";
    file << " <Style ss:ID=\"sCurrency\">\n";
    file << "  <NumberFormat ss:Format=\"$#,##0.00\"/>\n";
    file << " </Style>\n";
    file << " <Style ss:ID=\"sCurrencyBold\">\n";
    file << "  <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Bold=\"1\"/>\n";
    file << "  <NumberFormat ss:Format=\"$#,##0.00\"/>\n";
    file << " </Style>\n";
    file << "</Styles>\n";
    
    file << "<Worksheet ss:Name=\"Invoice\">\n";
    file << "<Table ss:ExpandedColumnCount=\"5\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">\n";
    file << "<Column ss:Width=\"150\"/>\n";
    file << "<Column ss:Width=\"100\"/>\n";
    file << "<Column ss:Width=\"100\"/>\n";
    
    file << "<Row ss:Height=\"20\">\n";
    file << "<Cell ss:MergeAcross=\"4\" ss:StyleID=\"sHeader\"><Data ss:Type=\"String\">INVOICE</Data></Cell>\n";
    file << "</Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">From:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_business_name->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">To:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_customer_name->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">Email:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_customer_email->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">Phone:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_customer_phone->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">Job Name:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_job_name->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">Description:</Data></Cell><Cell><Data ss:Type=\"String\">" << escape_xml(m_txt_job_description->GetValue().ToStdString()) << "</Data></Cell></Row>\n";
    file << "<Row><Cell ss:StyleID=\"sBold\"><Data ss:Type=\"String\">Date:</Data></Cell><Cell><Data ss:Type=\"String\">" << wxDateTime::Now().FormatDate() << "</Data></Cell></Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    file << "<Row ss:StyleID=\"sBold\">\n";
    file << "<Cell><Data ss:Type=\"String\">Item</Data></Cell>\n";
    file << "<Cell><Data ss:Type=\"String\">Quantity</Data></Cell>\n";
    file << "<Cell><Data ss:Type=\"String\">Unit Price</Data></Cell>\n";
    file << "<Cell><Data ss:Type=\"String\">Total</Data></Cell>\n";
    file << "</Row>\n";
    
    int total_parts = m_parts_per_plate->GetValue() * m_num_plates->GetValue();
    file << "<Row>\n";
    file << "<Cell><Data ss:Type=\"String\">3D Printed Parts</Data></Cell>\n";
    file << "<Cell><Data ss:Type=\"Number\">" << total_parts << "</Data></Cell>\n";
    file << "<Cell ss:StyleID=\"sCurrency\"><Data ss:Type=\"Number\">" << m_calc_final_price << "</Data></Cell>\n";
    file << "<Cell ss:StyleID=\"sCurrency\"><Data ss:Type=\"Number\">" << m_calc_total_job_cost << "</Data></Cell>\n";
    file << "</Row>\n";
    
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    file << "<Row ss:StyleID=\"sBold\"><Cell><Data ss:Type=\"String\">Material Breakdown</Data></Cell></Row>\n";
    for (const auto& data : m_filament_data) {
        file << "<Row>\n";
        file << "<Cell><Data ss:Type=\"String\">" << escape_xml(data.name) << " (" << escape_xml(data.color) << ")</Data></Cell>\n";
        file << "<Cell><Data ss:Type=\"String\">" << wxString::Format("%.2f g", data.weight_g) << "</Data></Cell>\n";
        file << "</Row>\n";
    }
    
    file << "</Table>\n";
    file << "</Worksheet>\n";
    
    file << "<Worksheet ss:Name=\"Internal Cost Breakdown\">\n";
    file << "<Table ss:ExpandedColumnCount=\"2\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">\n";
    file << "<Column ss:Width=\"200\"/>\n";
    file << "<Column ss:Width=\"100\"/>\n";
    
    file << "<Row ss:StyleID=\"sBold\"><Cell><Data ss:Type=\"String\">INTERNAL COST BREAKDOWN</Data></Cell></Row>\n";
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    
    auto add_cost_row = [&](const std::string& label, double val) {
        file << "<Row>\n";
        file << "<Cell><Data ss:Type=\"String\">" << label << "</Data></Cell>\n";
        file << "<Cell ss:StyleID=\"sCurrency\"><Data ss:Type=\"Number\">" << val << "</Data></Cell>\n";
        file << "</Row>\n";
    };
    
    add_cost_row("Material Cost", m_calc_material_cost);
    add_cost_row("Labor Cost", m_calc_labor_cost);
    add_cost_row("Machine Cost", m_calc_machine_cost);
    add_cost_row("Tooling Cost", m_calc_tooling_cost);
    add_cost_row("Post-Processing Cost", m_calc_postprocess_cost);
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    add_cost_row("Subtotal", m_calc_subtotal);
    add_cost_row("Failure Adjustment", m_calc_failure_adjustment);
    add_cost_row("Markup Amount", m_calc_markup_amount);
    file << "<Row><Cell><Data ss:Type=\"String\"></Data></Cell></Row>\n";
    add_cost_row("Total Job Cost", m_calc_total_job_cost);
    
    file << "</Table>\n";
    file << "</Worksheet>\n";
    
    file << "</Workbook>\n";
    
    output.Close();
    wxMessageBox(_L("Invoice exported successfully."), _L("Export Complete"), wxOK | wxICON_INFORMATION);
}

void InvoiceDialog::on_export_invoice(wxCommandEvent& event)
{
    wxFileDialog saveDialog(this, _L("Export Invoice"), "", "invoice.xls",
                            "Excel Files (*.xls)|*.xls",
                            wxFD_SAVE | wxFD_OVERWRITE_PROMPT);
    
    if (saveDialog.ShowModal() == wxID_CANCEL)
        return;
        
    export_to_excel(saveDialog.GetPath());
}

} // namespace GUI
} // namespace Slic3r
