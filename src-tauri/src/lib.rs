use serde::{Deserialize, Serialize};
use std::process::Command;
use std::os::windows::process::CommandExt;
use std::sync::Mutex;
use tauri::State;
use std::path::{Path, PathBuf};
use std::fs;
use std::time::{SystemTime, UNIX_EPOCH};

const CREATE_NO_WINDOW: u32 = 0x08000000;

#[derive(Default)]
struct AppState {
    current_engine: Mutex<String>,
}

#[derive(Serialize)]
struct EngineStatus {
    available: bool,
    engine: String,
    all_engines: Vec<String>,
}

#[tauri::command]
fn check_engines(state: State<'_, AppState>) -> Result<EngineStatus, String> {
    let mut all_engines = Vec::new();

    // Check Office COM
    let office_check = Command::new("reg")
        .args(["query", r#"HKCR\PowerPoint.Application"#, "/ve"])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    if let Ok(out) = office_check {
        if out.status.success() {
            all_engines.push("office".to_string());
        }
    }

    // Check WPS COM
    let wps_check = Command::new("reg")
        .args(["query", r#"HKCR\KWPP.Application"#, "/ve"])
        .creation_flags(CREATE_NO_WINDOW)
        .output();
    if let Ok(out) = wps_check {
        if out.status.success() {
            all_engines.push("wps".to_string());
        }
    }

    // Check LibreOffice Fallback
    let lo_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ];
    for p in lo_paths.iter() {
        if Path::new(p).exists() {
            all_engines.push("libreoffice".to_string());
            break;
        }
    }

    if all_engines.is_empty() {
        return Ok(EngineStatus {
            available: false,
            engine: "".to_string(),
            all_engines: vec![],
        });
    }

    let selected = all_engines[0].clone();
    *state.current_engine.lock().unwrap() = selected.clone();

    Ok(EngineStatus {
        available: true,
        engine: selected,
        all_engines,
    })
}

#[tauri::command]
fn set_engine(engine: String, state: State<'_, AppState>) -> Result<bool, String> {
    *state.current_engine.lock().unwrap() = engine;
    Ok(true)
}

#[tauri::command]
async fn convert_pptx(path: String, state: State<'_, AppState>) -> Result<String, String> {
    let engine = state.current_engine.lock().unwrap().clone();
    let current_dir = std::env::temp_dir();
    let input_path = PathBuf::from(&path);
    if !input_path.exists() {
        return Err("输入文件不存在".to_string());
    }

    let file_stem = input_path.file_stem().unwrap().to_string_lossy();
    let parent = input_path.parent().unwrap_or(Path::new(""));
    let pdf_path = parent.join(format!("{}.pdf", file_stem));
    let pdf_path_str = pdf_path.to_string_lossy().to_string();

    match engine.as_str() {
        "office" => {
            let script = format!(
                r#"
$ErrorActionPreference = "Stop"
$ppt = $null; $pres = $null
try {{
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open('{}', 2, 0, 0)
    $pres.SaveAs('{}', [int]32)
    $pres.Close()
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}} finally {{
    if ($pres) {{ try {{ [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) }} catch {{}} }}
    if ($ppt) {{ try {{ $ppt.Quit() }} catch {{}}; try {{ [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) }} catch {{}} }}
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}}
                "#,
                path.replace("'", "''"),
                pdf_path_str.replace("'", "''")
            );
            
            let timestamp = SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis();
            let script_path = current_dir.join(format!("ppt2pdf_office_{}.ps1", timestamp));
            // powershell needs UTF-8 WITH BOM to correctly parse paths handling spaces and non-ascii
            let mut content = vec![0xEF, 0xBB, 0xBF];
            content.extend_from_slice(script.as_bytes());
            fs::write(&script_path, content).map_err(|e| e.to_string())?;

            let output = Command::new("powershell")
                .creation_flags(CREATE_NO_WINDOW)
                .args(["-ExecutionPolicy", "Bypass", "-NoProfile", "-WindowStyle", "Hidden", "-File", script_path.to_str().unwrap()])
                .output()
                .map_err(|e| e.to_string())?;

            let _ = fs::remove_file(&script_path);

            if !output.status.success() {
                return Err(String::from_utf8_lossy(&output.stderr).to_string());
            }

            if !pdf_path.exists() {
                let double_pdf = format!("{}.pdf", pdf_path_str);
                if Path::new(&double_pdf).exists() {
                    let _ = fs::rename(&double_pdf, &pdf_path_str);
                }
            }
            if !pdf_path.exists() {
                let no_ext = pdf_path_str.trim_end_matches(".pdf");
                if Path::new(no_ext).exists() {
                    let _ = fs::rename(no_ext, &pdf_path_str);
                }
            }

            if pdf_path.exists() {
                Ok(pdf_path_str)
            } else {
                Err("Office生成PDF失败".to_string())
            }
        },
        "wps" => {
            let script = format!(
                r#"
$ErrorActionPreference = "Stop"
$wpp = $null; $pres = $null
try {{
    $wpp = New-Object -ComObject KWPP.Application
    $pres = $wpp.Presentations.Open('{}', $true, $false, $false)
    $pres.SaveAs('{}', [int]32)
    $pres.Close()
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}} finally {{
    if ($pres) {{ try {{ [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) }} catch {{}} }}
    if ($wpp) {{ try {{ $wpp.Quit() }} catch {{}}; try {{ [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wpp) }} catch {{}} }}
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}}
                "#,
                path.replace("'", "''"),
                pdf_path_str.replace("'", "''")
            );
            
            let timestamp = SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_millis();
            let script_path = current_dir.join(format!("ppt2pdf_wps_{}.ps1", timestamp));
            let mut content = vec![0xEF, 0xBB, 0xBF];
            content.extend_from_slice(script.as_bytes());
            fs::write(&script_path, content).map_err(|e| e.to_string())?;

            let output = Command::new("powershell")
                .creation_flags(CREATE_NO_WINDOW)
                .args(["-ExecutionPolicy", "Bypass", "-NoProfile", "-WindowStyle", "Hidden", "-File", script_path.to_str().unwrap()])
                .output()
                .map_err(|e| e.to_string())?;

            let _ = fs::remove_file(&script_path);

            if !output.status.success() {
                return Err(String::from_utf8_lossy(&output.stderr).to_string());
            }

            if !pdf_path.exists() {
                let double_pdf = format!("{}.pdf", pdf_path_str);
                if Path::new(&double_pdf).exists() {
                    let _ = fs::rename(&double_pdf, &pdf_path_str);
                }
            }
            if !pdf_path.exists() {
                let no_ext = pdf_path_str.trim_end_matches(".pdf");
                if Path::new(no_ext).exists() {
                    let _ = fs::rename(no_ext, &pdf_path_str);
                }
            }

            if pdf_path.exists() {
                Ok(pdf_path_str)
            } else {
                Err("WPS生成PDF失败".to_string())
            }
        },
        "libreoffice" => {
            let mut soffice = String::new();
            let lo_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ];
            for p in lo_paths.iter() {
                if Path::new(p).exists() {
                    soffice = p.to_string();
                    break;
                }
            }
            if soffice.is_empty() {
                return Err("LibreOffice 未安装".to_string());
            }

            let output = Command::new(soffice)
                .creation_flags(CREATE_NO_WINDOW)
                .args([
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", parent.to_str().unwrap(),
                    &path
                ])
                .output()
                .map_err(|e| e.to_string())?;

            if !output.status.success() {
                return Err(String::from_utf8_lossy(&output.stderr).to_string());
            }

            if pdf_path.exists() {
                Ok(pdf_path_str)
            } else {
                Err("生成 PDF 失败".to_string())
            }
        },
        _ => Err(format!("未知引擎: {}", engine)),
    }
}

#[tauri::command]
fn open_file(path: String) -> Result<(), String> {
    Command::new("cmd")
        .creation_flags(CREATE_NO_WINDOW)
        .args(["/C", "start", "", &path])
        .spawn()
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn open_folder(path: String) -> Result<(), String> {
    let p = PathBuf::from(&path);
    if let Some(parent) = p.parent() {
        Command::new("cmd")
            .creation_flags(CREATE_NO_WINDOW)
            .args(["/C", "start", "", parent.to_str().unwrap()])
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

#[tauri::command]
async fn select_files() -> Result<Vec<String>, String> {
    if let Some(files) = rfd::AsyncFileDialog::new()
        .add_filter("PowerPoint", &["ppt", "pptx", "pps", "ppsx"])
        .pick_files()
        .await
    {
        Ok(files.into_iter().map(|f| f.path().to_string_lossy().into_owned()).collect())
    } else {
        Ok(vec![])
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(AppState::default())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            check_engines,
            set_engine,
            convert_pptx,
            open_file,
            open_folder,
            select_files
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
