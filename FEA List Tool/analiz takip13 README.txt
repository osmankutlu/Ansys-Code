ANALYSIS TRACKER - SETUP AND USE
================================

SETUP (once, done by one person / the administrator)
1. Copy this whole folder (Analysis Tracker Launcher 4.xlsm + the AnalizTakip folder) to a network folder that EVERYONE can write to.
   Example: \\server\Engineering\AnalysisTracker\
2. Everyone opens "Analysis Tracker Launcher 4.xlsm" from that same shared folder (do not copy it to a local PC).
3. In Excel, click "Enable Content" / "Enable Macros" in the yellow bar.
4. Press "OPEN PROGRAM". The program opens in its own window.

Data is stored in AnalizTakip\Veri (one file per analysis). Please back this folder up.
If an error occurs, send AnalizTakip\AnalizTakip.hata.log (the "Error Log" button in Excel opens it).

USING THE PROGRAM
- New Analysis (green button, bottom left of the list): creates a new request.
  Part names and materials can have several entries: type one and press Add (or Enter).
  Model Type, Material Model (Elastic / Elastoplastic) and Alpha (On / Off) can have several selections.
- Model Role: Standalone / Master Model / Sub-Analysis. The "Linked Master Model" list becomes active only for Sub-Analysis;
  choosing a master model copies the model information, and the record is filed under that master model.
- Status, Analyst Notes, Model File Path, Results Path and Completion Date are not asked for new requests:
  open the record afterwards (double-click / Edit) and let the analyst fill them in.
- "Master Model Tree" tab: shows which analyses belong to which master model.
- "Summary" tab: open / overdue counts and distribution by person and status.
- Several people can work at the same time; if a record was changed by someone else you are warned and your save is blocked.
- Deleted records are moved to Veri\Silinenler; IDs are never reused.
- Lists (status, material, condition, people ...): Settings > Edit lists.

USERS, ROLES AND PERMISSIONS (new)
- First start on each PC: the program asks for your name and role (Leader / Design Engineer / Structural Analysis Engineer).
  You are tied to your Windows account (DOMAIN\user); next time you are recognised automatically.
  The very first person who starts the program becomes the administrator, so the administrator should start it first.
- Leader and Structural Analysis Engineer can edit every record. Design Engineers can edit only their own requests
  (request fields only; status, analyst notes, file paths and completion date are filled in by analysts), and can delete their own
  request only while its status is still Pending. Everything else is read-only for them. Closed (Completed/Cancelled) requests are read-only for designers.
- Only the administrator can change a name or role (Admin > Manage users). If a user is removed there, the program closes on that PC.
- Extra administrators: Admin > Manage users > Admin on/off, or list Windows accounts (DOMAIN\user, one per line) in AdminAccounts.txt.

NOTIFICATIONS (new)
- Design Engineer creates or changes a request: the assigned analyst is notified; if no analyst is assigned, all Structural Analysis Engineers are notified.
- An analyst / leader changes a record: the requester is notified. A newly assigned analyst is notified.
- Leaders never receive notifications. Nobody is notified about their own actions.
- A balloon appears near the clock and the toolbar button shows the unread count ("Notifications (3)"). Notifications work while the program is open (minimised is fine).
  Old notifications (30+ days) are removed automatically.

DUPLICATING
- Model Type, Material Model and Alpha are single-choice. To request the same analysis with another option (e.g. Alpha On, then Alpha Off),
  use "Save + Duplicate" in the request window, or select a record and click Duplicate in the toolbar. The copy opens with the same data;
  change the option that differs and save.

SAMPLE DATA
- On the very first start (no records) the administrator is offered sample data: 8 master models with 16 sub-analyses and 4 standalone analyses,
  using the registered users as requesters/analysts. Admin > Add sample data adds it again; delete the sample records from the list when done.

ADMIN MENU
- Manage users, Restore deleted records, Reassign analyst (hand over open work, e.g. holidays), Back up data now (keeps last 10 in Veri\Yedek),
  Change data folder, Add sample data.

NOTE ON OLDER DATA
The program is now fully in English. Records created with the earlier Turkish version are read and converted automatically
(status, priority, role, etc.); text you typed yourself (part names, notes) stays as it was. Sample data from the earlier test
versions can simply be deleted from the program.

REQUIREMENTS
- Windows + Excel (macro-enabled .xlsm) + Windows PowerShell 5.1 (built into Windows 10/11). No Python or other installation needed.
- If company security policy blocks macros or PowerShell, ask IT to allow running AnalizTakip.ps1.
- To open the program without Excel: powershell -STA -ExecutionPolicy Bypass -File AnalizTakip\AnalizTakip.ps1
