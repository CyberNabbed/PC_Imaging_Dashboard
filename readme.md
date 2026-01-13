# Device Enrollment Dashboard

This tool was built to help IT staff monitor the realtime status of device enrollments (Entra Join + Intune) without manually digging through raw log files on the server.

**What it does:**
* **Live Monitoring:** Reads logs directly from the SCCM network share to display status. Rows turn Green for success, Red for failure, and Yellow while enrolling.
* **Data Fetching:** Automatically pulls Serial Numbers in the background using PowerShell so you don't have to look them up.
* **Quick Add:** Use "Browse Recent" to quickly see and add machines imaged in the last 24 hours.
* **Reporting:** Click "Report to O'Leary" to generate a pre-formatted Outlook email draft for failed machines, including their log details and duration.

**How to use:**
1. Ensure you have Python installed and are connected to the network (access to `\\sccm...` is required).
2. Run the script.
3. Type a hostname or use the "Browse Recent" button to start tracking.
4. The dashboard auto-refreshes every 5 minutes.
