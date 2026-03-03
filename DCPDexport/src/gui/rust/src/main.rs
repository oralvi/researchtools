// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod commands;
mod models;

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            commands::analyze_dcpd_data,
            commands::process_file,
            commands::get_status,
            commands::get_examples,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
