#ifndef slic3r_GUI_InvoiceDialog_hpp_
#define slic3r_GUI_InvoiceDialog_hpp_

#include <wx/wx.h>
#include <wx/spinctrl.h>
#include <wx/grid.h>
#include <vector>
#include <map>
#include <string>
#include "GUI_Utils.hpp"
#include "wxExtensions.hpp"

namespace Slic3r {

// Forward declarations
struct PrintStatistics;

namespace GUI {

// Structure to hold per-filament data
struct FilamentData {
    size_t extruder_id;
    std::string name;
    std::string color;
    double weight_g;        // grams used
    double cost_per_kg;     // $/kg from profile or user-set
    double calculated_cost; // computed cost
};

// Structure to hold job profile
struct JobProfile {
    std::string job_name;
    std::string customer_name;
    std::string customer_email;
    std::string customer_phone;
    std::string job_description;
    
    // Job parameters
    int parts_per_plate;
    int num_plates;
    double failure_rate;
    
    // Labor
    double labor_rate;
    double prep_time;
    double setup_time;
    double finishing_per_part;
    double finishing_per_plate;
    
    // Machine
    double printer_cost;
    double printer_lifespan;
    double maintenance_cost;
    double power_watts;
    double electricity_cost;
    
    // Tooling
    double bed_cost;
    double bed_lifespan;
    double nozzle_cost;
    double nozzle_lifespan_kg;
    
    // Post-Processing
    double solvent_cost;
    double solving_time;
    double tank_power;
    double finishing_materials;
    
    // Markup
    double markup_percent;
    
    // Per-filament cost overrides (extruder_id -> cost_per_kg)
    std::map<size_t, double> filament_costs;
};

class InvoiceDialog : public DPIDialog
{
public:
    InvoiceDialog(wxWindow* parent, const PrintStatistics* stats);
    ~InvoiceDialog() = default;

protected:
    void on_dpi_changed(const wxRect& suggested_rect) override;

private:
    void build_ui();
    void build_customer_info_tab(wxNotebook* notebook);
    void build_job_info_tab(wxNotebook* notebook);
    void build_materials_tab(wxNotebook* notebook);
    void build_labor_tab(wxNotebook* notebook);
    void build_machine_tab(wxNotebook* notebook);
    void build_tooling_tab(wxNotebook* notebook);
    void build_postprocess_tab(wxNotebook* notebook);
    void build_markup_tab(wxNotebook* notebook);
    void build_results_tab(wxNotebook* notebook);
    
    void load_global_settings();
    void save_global_settings();
    void load_job_profile(const std::string& profile_name);
    void save_job_profile(const std::string& profile_name);
    void delete_job_profile(const std::string& profile_name);
    std::vector<std::string> get_saved_job_profiles();
    void refresh_job_profiles_combo();
    
    void populate_filament_data();
    void update_materials_grid();
    void calculate_costs();
    
    void on_calculate(wxCommandEvent& event);
    void on_save_job(wxCommandEvent& event);
    void on_load_job(wxCommandEvent& event);
    void on_delete_job(wxCommandEvent& event);
    void on_export_invoice(wxCommandEvent& event);
    void on_filament_cost_changed(wxGridEvent& event);
    
    void export_to_excel(const wxString& path);
    std::string escape_xml(const std::string& str) const;
    
    wxString format_time(const std::string& time_str) const;
    double parse_time_to_hours(const std::string& time_str) const;

    // Print statistics from slicer
    const PrintStatistics* m_stats;
    
    // Filament data
    std::vector<FilamentData> m_filament_data;
    
    // Current job profile
    JobProfile m_current_job;

    // === Customer/Vendor Info Controls ===
    wxTextCtrl* m_txt_business_name;
    wxTextCtrl* m_txt_customer_name;
    wxTextCtrl* m_txt_customer_email;
    wxTextCtrl* m_txt_customer_phone;
    wxTextCtrl* m_txt_job_name;
    wxTextCtrl* m_txt_job_description;
    
    // Job profile management
    wxComboBox* m_combo_job_profiles;

    // === Job Parameters Controls ===
    wxSpinCtrl* m_parts_per_plate;
    wxSpinCtrl* m_num_plates;
    wxSpinCtrlDouble* m_failure_rate;

    // === Materials Controls ===
    wxGrid* m_materials_grid;
    wxStaticText* m_lbl_total_material_cost;

    // === Labor Controls ===
    wxSpinCtrlDouble* m_labor_rate;
    wxSpinCtrlDouble* m_prep_time;
    wxSpinCtrlDouble* m_setup_time;
    wxSpinCtrlDouble* m_finishing_per_part;
    wxSpinCtrlDouble* m_finishing_per_plate;

    // === Machine Controls ===
    wxSpinCtrlDouble* m_printer_cost;
    wxSpinCtrlDouble* m_printer_lifespan;
    wxSpinCtrlDouble* m_maintenance_cost;
    wxSpinCtrlDouble* m_power_watts;
    wxSpinCtrlDouble* m_electricity_cost;

    // === Tooling Controls ===
    wxSpinCtrlDouble* m_bed_cost;
    wxSpinCtrlDouble* m_bed_lifespan;
    wxSpinCtrlDouble* m_nozzle_cost;
    wxSpinCtrlDouble* m_nozzle_lifespan_kg;

    // === Post-Processing Controls ===
    wxSpinCtrlDouble* m_solvent_cost;
    wxSpinCtrlDouble* m_solving_time;
    wxSpinCtrlDouble* m_tank_power;
    wxSpinCtrlDouble* m_finishing_materials;

    // === Markup Controls ===
    wxSpinCtrlDouble* m_markup_percent;

    // === Slicer Data Display ===
    wxStaticText* m_lbl_print_time;
    wxStaticText* m_lbl_total_weight;

    // === Results Display ===
    wxStaticText* m_lbl_material_cost;
    wxStaticText* m_lbl_labor_cost;
    wxStaticText* m_lbl_machine_cost;
    wxStaticText* m_lbl_tooling_cost;
    wxStaticText* m_lbl_postprocess_cost;
    wxStaticText* m_lbl_subtotal;
    wxStaticText* m_lbl_failure_adjustment;
    wxStaticText* m_lbl_cost_per_part;
    wxStaticText* m_lbl_markup_amount;
    wxStaticText* m_lbl_final_price;
    wxStaticText* m_lbl_total_job_cost;

    // === Buttons ===
    wxButton* m_btn_calculate;
    wxButton* m_btn_save_job;
    wxButton* m_btn_load_job;
    wxButton* m_btn_delete_job;
    wxButton* m_btn_export;
    wxButton* m_btn_close;
    
    // Cached calculation results for export
    double m_calc_material_cost = 0.0;
    double m_calc_labor_cost = 0.0;
    double m_calc_machine_cost = 0.0;
    double m_calc_tooling_cost = 0.0;
    double m_calc_postprocess_cost = 0.0;
    double m_calc_subtotal = 0.0;
    double m_calc_failure_adjustment = 0.0;
    double m_calc_cost_per_part = 0.0;
    double m_calc_markup_amount = 0.0;
    double m_calc_final_price = 0.0;
    double m_calc_total_job_cost = 0.0;
    double m_calc_print_time_hours = 0.0;
};

} // namespace GUI
} // namespace Slic3r

#endif // slic3r_GUI_InvoiceDialog_hpp_
