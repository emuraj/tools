# flatten_files_streamlit.py
"""
Purpose:
    Provide a simple Streamlit UI for flattening files from three folder
    hierarchies into a single output folder.

What this file does:
    - Presents a browser-based UI using Streamlit
    - Provides Browse buttons using Windows PowerShell dialogs via subprocess
    - Recursively scans 3 input directories and all subdirectories
    - Copies every file into 1 output directory
    - Renames duplicate filenames by appending _1, _2, etc.
    - Writes a TXT report listing files that were renamed
    - Writes a professional boilerplate message if no renaming was required
    - Constrains the UI to a centered, document-like width

Place in the larger scheme:
    Use this as a lightweight local utility when you want a simple UI instead
    of editing path variables directly in a script.
"""

from __future__ import annotations

from pathlib import Path
import shutil
import subprocess

import streamlit as st


NO_DUPLICATES_MESSAGE = (
    "Processing completed successfully. "
    "No duplicate filenames were detected, and no renaming was necessary."
)


def run_powershell(ps_script: str) -> str:
    """
    Run a PowerShell script and return stdout as stripped text.
    Returns an empty string if the dialog is cancelled or PowerShell errors.
    """
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-STA", "-Command", ps_script],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            return ""
        return result.stdout.strip()
    except Exception:
        return ""


def escape_ps_single_quoted(value: str) -> str:
    """
    Escape a string for safe insertion into a PowerShell single-quoted string.
    """
    return value.replace("'", "''")


def choose_directory(initial_value: str = "") -> str:
    """
    Open a Windows folder picker using PowerShell and return the selected path.
    Returns the original value if the user cancels.
    """
    initial_value = initial_value.strip()

    if initial_value and Path(initial_value).exists():
        selected_path = escape_ps_single_quoted(initial_value)
        ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.SelectedPath = '{selected_path}'
$dialog.ShowNewFolderButton = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $dialog.SelectedPath
}}
"""
    else:
        ps_script = """
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.ShowNewFolderButton = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.SelectedPath
}
"""

    selected = run_powershell(ps_script)
    return selected if selected else initial_value


def choose_save_file(initial_value: str = "") -> str:
    """
    Open a Windows save-file picker using PowerShell and return the selected path.
    Returns the original value if the user cancels.
    """
    initial_value = initial_value.strip()
    initial_dir = ""
    initial_file = "renamed_files_report.txt"

    if initial_value:
        p = Path(initial_value)
        if p.parent.exists():
            initial_dir = str(p.parent)
        if p.name:
            initial_file = p.name

    initial_dir = escape_ps_single_quoted(initial_dir)
    initial_file = escape_ps_single_quoted(initial_file)

    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$dialog = New-Object System.Windows.Forms.SaveFileDialog
$dialog.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
$dialog.DefaultExt = 'txt'
$dialog.AddExtension = $true
$dialog.FileName = '{initial_file}'
if ('{initial_dir}' -ne '') {{
    $dialog.InitialDirectory = '{initial_dir}'
}}
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    Write-Output $dialog.FileName
}}
"""

    selected = run_powershell(ps_script)
    return selected if selected else initial_value


def get_unique_destination_path(destination_dir: Path, filename: str) -> tuple[Path, bool]:
    """
    Return a unique file path inside destination_dir.

    Returns:
        (final_path, was_renamed)
    """
    candidate = destination_dir / filename

    if not candidate.exists():
        return candidate, False

    stem = candidate.stem
    suffix = candidate.suffix
    counter = 1

    while True:
        new_candidate = destination_dir / f"{stem}_{counter}{suffix}"
        if not new_candidate.exists():
            return new_candidate, True
        counter += 1


def copy_all_files_to_single_folder(
    input_dirs: list[Path],
    output_dir: Path,
) -> tuple[list[str], int, int]:
    """
    Recursively copy all files from input_dirs into output_dir.

    Returns:
        renamed_files_report: list of report lines for renamed files
        total_copied: total number of files copied
        total_skipped: total number of invalid input roots skipped
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    renamed_files_report: list[str] = []
    total_copied = 0
    total_skipped = 0

    for input_dir in input_dirs:
        if not input_dir.exists() or not input_dir.is_dir():
            total_skipped += 1
            continue

        for item in input_dir.rglob("*"):
            if not item.is_file():
                continue

            destination_path, was_renamed = get_unique_destination_path(output_dir, item.name)
            shutil.copy2(item, destination_path)
            total_copied += 1

            if was_renamed:
                renamed_files_report.append(
                    f"Original source: {item}\n"
                    f"Copied as:       {destination_path.name}\n"
                )

    return renamed_files_report, total_copied, total_skipped


def write_report(report_txt_path: Path, renamed_files_report: list[str]) -> None:
    """
    Write the rename report to a text file.
    """
    report_txt_path.parent.mkdir(parents=True, exist_ok=True)

    with report_txt_path.open("w", encoding="utf-8") as report_file:
        if renamed_files_report:
            report_file.write("Files renamed due to duplicate filenames\n")
            report_file.write("=" * 50 + "\n\n")
            report_file.write("\n".join(renamed_files_report))
        else:
            report_file.write(NO_DUPLICATES_MESSAGE + "\n")


def init_session_state() -> None:
    """
    Initialize Streamlit session state keys.
    """
    defaults = {
        "input_dir_1": "",
        "input_dir_2": "",
        "input_dir_3": "",
        "output_dir": "",
        "report_txt_path": "",
        "input_dir_1_display": "",
        "input_dir_2_display": "",
        "input_dir_3_display": "",
        "output_dir_display": "",
        "report_txt_path_display": "",
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def sync_display_to_value(value_key: str, display_key: str) -> None:
    """
    Copy the widget display value into the stored value.
    """
    st.session_state[value_key] = st.session_state[display_key]


def browse_directory_callback(value_key: str, display_key: str) -> None:
    """
    Browse for a directory and update both stored and display values.
    """
    selected = choose_directory(st.session_state[value_key])
    if selected:
        st.session_state[value_key] = selected
        st.session_state[display_key] = selected


def browse_file_callback(value_key: str, display_key: str) -> None:
    """
    Browse for a save-file path and update both stored and display values.
    """
    selected = choose_save_file(st.session_state[value_key])
    if selected:
        st.session_state[value_key] = selected
        st.session_state[display_key] = selected


def render_directory_row(label: str, value_key: str, display_key: str) -> None:
    """
    Render a text input plus Browse button for a directory.
    """
    col1, col2 = st.columns([7, 1])

    with col1:
        st.text_input(
            label,
            key=display_key,
            on_change=sync_display_to_value,
            args=(value_key, display_key),
        )

    with col2:
        st.write("")
        st.button(
            "Browse",
            key=f"{value_key}_browse",
            on_click=browse_directory_callback,
            args=(value_key, display_key),
        )


def render_file_row(label: str, value_key: str, display_key: str) -> None:
    """
    Render a text input plus Browse button for a file save path.
    """
    col1, col2 = st.columns([7, 1])

    with col1:
        st.text_input(
            label,
            key=display_key,
            on_change=sync_display_to_value,
            args=(value_key, display_key),
        )

    with col2:
        st.write("")
        st.button(
            "Browse",
            key=f"{value_key}_browse",
            on_click=browse_file_callback,
            args=(value_key, display_key),
        )


def validate_required_fields() -> list[str]:
    """
    Validate required UI fields.
    """
    missing_fields = []

    field_map = {
        "Input Folder 1": "input_dir_1",
        "Input Folder 2": "input_dir_2",
        "Input Folder 3": "input_dir_3",
        "Output Folder": "output_dir",
        "Report TXT Path": "report_txt_path",
    }

    for display_name, state_key in field_map.items():
        if not st.session_state[state_key].strip():
            missing_fields.append(display_name)

    return missing_fields


def validate_paths() -> list[str]:
    """
    Validate entered paths for existence and correctness.
    """
    errors: list[str] = []

    input_keys = [
        ("Input Folder 1", "input_dir_1"),
        ("Input Folder 2", "input_dir_2"),
        ("Input Folder 3", "input_dir_3"),
    ]

    for display_name, state_key in input_keys:
        value = st.session_state[state_key].strip()
        if value:
            p = Path(value)
            if not p.exists():
                errors.append(f"{display_name} does not exist.")
            elif not p.is_dir():
                errors.append(f"{display_name} is not a folder.")

    output_value = st.session_state["output_dir"].strip()
    if output_value:
        output_path = Path(output_value)
        parent = output_path.parent
        if output_path.exists() and not output_path.is_dir():
            errors.append("Output Folder exists but is not a folder.")
        elif not output_path.exists() and not parent.exists():
            errors.append("Output Folder parent path does not exist.")

    report_value = st.session_state["report_txt_path"].strip()
    if report_value:
        report_path = Path(report_value)
        if report_path.suffix.lower() != ".txt":
            errors.append("Report TXT Path must end with .txt.")
        if not report_path.parent.exists():
            errors.append("Report TXT Path parent folder does not exist.")

    return errors


def main() -> None:
    """
    Run the Streamlit application.
    """
    st.set_page_config(page_title="Flatten Files Utility", layout="wide")

    st.markdown(
        """
        <style>
            .block-container {
                max-width: 8.5in;
                margin-left: auto;
                margin-right: auto;
                padding-top: 2rem;
                padding-bottom: 2rem;
                padding-left: 0.75in;
                padding-right: 0.75in;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    init_session_state()

    st.title("Flatten Files Utility")
    st.write(
        "Copy every file from three folder hierarchies into one output folder. "
        "Duplicate filenames are automatically renamed, and a TXT report is generated."
    )

    st.subheader("Locations")

    render_directory_row("Input Folder 1", "input_dir_1", "input_dir_1_display")
    render_directory_row("Input Folder 2", "input_dir_2", "input_dir_2_display")
    render_directory_row("Input Folder 3", "input_dir_3", "input_dir_3_display")
    render_directory_row("Output Folder", "output_dir", "output_dir_display")
    render_file_row("Report TXT Path", "report_txt_path", "report_txt_path_display")

    st.divider()

    if st.button("Execute", type="primary", use_container_width=True):
        missing_fields = validate_required_fields()
        if missing_fields:
            for field_name in missing_fields:
                st.error(f"{field_name} is required.")
            st.stop()

        path_errors = validate_paths()
        if path_errors:
            for error in path_errors:
                st.error(error)
            st.stop()

        input_dirs = [
            Path(st.session_state["input_dir_1"]),
            Path(st.session_state["input_dir_2"]),
            Path(st.session_state["input_dir_3"]),
        ]
        output_dir = Path(st.session_state["output_dir"])
        report_txt_path = Path(st.session_state["report_txt_path"])

        with st.spinner("Processing files..."):
            renamed_files_report, total_copied, total_skipped = copy_all_files_to_single_folder(
                input_dirs=input_dirs,
                output_dir=output_dir,
            )
            write_report(report_txt_path, renamed_files_report)

        st.success("Processing completed successfully.")
        st.write(f"Files copied: {total_copied}")
        st.write(f"Invalid input roots skipped: {total_skipped}")
        st.write(f"Output folder: {output_dir}")
        st.write(f"Report path: {report_txt_path}")

        if renamed_files_report:
            st.info(f"Duplicate filenames renamed: {len(renamed_files_report)}")
        else:
            st.info(NO_DUPLICATES_MESSAGE)


if __name__ == "__main__":
    main()