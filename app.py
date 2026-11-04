import sys, os, sqlite3, csv, shutil, ctypes, time, json
from pathlib import Path
from datetime import datetime, timedelta, date
import psutil
from PySide6.QtCore import Qt, QTimer, QDate, QTime, Signal, QObject
from PySide6.QtGui import QAction, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QGridLayout,QLabel,QPushButton,QLineEdit,QComboBox,QTableWidget,QTableWidgetItem,QHeaderView,QStackedWidget,QFrame,QDialog,QFormLayout,QDateEdit,QTimeEdit,QSpinBox,QCheckBox,QMessageBox,QFileDialog,QTabWidget,QDoubleSpinBox,QListWidget,QInputDialog,QSystemTrayIcon,QMenu,QStyle,QTextEdit)
from PySide6.QtCharts import QChart,QChartView,QPieSeries,QBarSeries,QBarSet,QBarCategoryAxis,QValueAxis
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
try: import keyboard
except Exception: keyboard=None

APP='ChronoFlow'; BASE=(Path.home()/'Library'/'Application Support'/APP) if sys.platform=='darwin' else Path(os.getenv('LOCALAPPDATA',Path.home()))/APP; BASE.mkdir(parents=True,exist_ok=True); DB=BASE/'chronoflow.db'; DEFAULT_BACKUP_DIR=BASE/'Backups'; DEFAULT_BACKUP_DIR.mkdir(parents=True,exist_ok=True)

def nowiso(): return datetime.now().isoformat(timespec='seconds')
def durfmt(sec):
    sec=max(0,int(sec)); h,rem=divmod(sec,3600); m,s=divmod(rem,60); return f'{h:02d}:{m:02d}:{s:02d}'
def parse_dt(s): return datetime.fromisoformat(s) if s else None

def fg_info():
    if sys.platform=='darwin':
        try:
            from AppKit import NSWorkspace
            import Quartz
            app=NSWorkspace.sharedWorkspace().frontmostApplication()
            name=app.localizedName() if app else 'Unknown'
            title=''
            if app:
                for win in Quartz.CGWindowListCopyWindowInfo(Quartz.kCGWindowListOptionOnScreenOnly,Quartz.kCGNullWindowID):
                    if win.get('kCGWindowOwnerPID')==app.processIdentifier() and win.get('kCGWindowLayer',1)==0:
                        title=win.get('kCGWindowName','') or ''
                        if title:break
            return name,title
        except Exception:return ('Unknown','')
    if os.name!='nt': return ('Development environment','')
    try:
        user32=ctypes.windll.user32; hwnd=user32.GetForegroundWindow(); n=user32.GetWindowTextLengthW(hwnd); buf=ctypes.create_unicode_buffer(n+1); user32.GetWindowTextW(hwnd,buf,n+1)
        pid=ctypes.c_ulong(); user32.GetWindowThreadProcessId(hwnd,ctypes.byref(pid)); name=psutil.Process(pid.value).name()
        return name,buf.value
    except Exception: return ('Unknown','')
def idle_seconds():
    if sys.platform=='darwin':
        try:
            import Quartz
            return float(Quartz.CGEventSourceSecondsSinceLastEventType(Quartz.kCGEventSourceStateCombinedSessionState,Quartz.kCGAnyInputEventType))
        except Exception:return 0
    if os.name!='nt': return 0
    class LASTINPUTINFO(ctypes.Structure): _fields_=[('cbSize',ctypes.c_uint),('dwTime',ctypes.c_uint)]
    li=LASTINPUTINFO(); li.cbSize=ctypes.sizeof(li); ctypes.windll.user32.GetLastInputInfo(ctypes.byref(li)); return (ctypes.windll.kernel32.GetTickCount()-li.dwTime)/1000

class Store:
    def __init__(self):
        self.c=sqlite3.connect(DB); self.c.row_factory=sqlite3.Row; self.init()
    def init(self):
        self.c.executescript('''
        CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY,name TEXT UNIQUE,color TEXT DEFAULT '#5B8DEF',rate REAL DEFAULT 0,estimate REAL DEFAULT 0,archived INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS tasks(id INTEGER PRIMARY KEY,project_id INTEGER,name TEXT,UNIQUE(project_id,name));
        CREATE TABLE IF NOT EXISTS tags(id INTEGER PRIMARY KEY,name TEXT UNIQUE);
        CREATE TABLE IF NOT EXISTS entries(id INTEGER PRIMARY KEY,description TEXT,project_id INTEGER,task_id INTEGER,start TEXT,end TEXT,duration INTEGER DEFAULT 0,tags TEXT DEFAULT '',favorite INTEGER DEFAULT 0,is_break INTEGER DEFAULT 0,billable INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS activity(id INTEGER PRIMARY KEY,app TEXT,title TEXT,start TEXT,end TEXT,duration INTEGER,idle INTEGER DEFAULT 0,converted INTEGER DEFAULT 0);
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
        '''); self.c.commit()
        if not self.one('SELECT id FROM projects LIMIT 1'): self.exec("INSERT INTO projects(name,color) VALUES('Personal','#5B8DEF')")
    def rows(self,q,p=()): return self.c.execute(q,p).fetchall()
    def one(self,q,p=()): return self.c.execute(q,p).fetchone()
    def exec(self,q,p=()): cur=self.c.execute(q,p); self.c.commit(); return cur.lastrowid
    def get(self,k,d=None): r=self.one('SELECT value FROM settings WHERE key=?',(k,)); return r['value'] if r else d
    def set(self,k,v): self.exec('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',(k,str(v)))

class EntryDialog(QDialog):
    def __init__(self,store,parent=None,row=None):
        super().__init__(parent); self.s=store; self.row=row; self.setWindowTitle('Time entry'); f=QFormLayout(self)
        self.desc=QLineEdit(); self.proj=QComboBox(); self.task=QComboBox(); self.tags=QLineEdit(); self.day=QDateEdit(QDate.currentDate()); self.day.setCalendarPopup(True); self.st=QTimeEdit(QTime.currentTime()); self.en=QTimeEdit(QTime.currentTime()); self.bill=QCheckBox('Billable')
        for r in self.s.rows('SELECT id,name FROM projects WHERE archived=0 ORDER BY name'): self.proj.addItem(r['name'],r['id'])
        self.proj.currentIndexChanged.connect(self.tasks); self.tasks()
        for w,n in [(self.desc,'Description'),(self.proj,'Project'),(self.task,'Task'),(self.tags,'Tags (comma separated)'),(self.day,'Date'),(self.st,'Start'),(self.en,'End')]: f.addRow(n,w)
        f.addRow('',self.bill); b=QPushButton('Save'); b.clicked.connect(self.accept); f.addRow(b)
        if row:
            self.desc.setText(row['description'] or ''); i=self.proj.findData(row['project_id']); self.proj.setCurrentIndex(max(i,0)); self.tasks(); j=self.task.findData(row['task_id']); self.task.setCurrentIndex(max(j,0)); self.tags.setText(row['tags'] or '')
            a=parse_dt(row['start']); z=parse_dt(row['end']) or a+timedelta(seconds=row['duration']); self.day.setDate(QDate(a.year,a.month,a.day)); self.st.setTime(QTime(a.hour,a.minute,a.second)); self.en.setTime(QTime(z.hour,z.minute,z.second)); self.bill.setChecked(bool(row['billable']))
    def tasks(self):
        old=self.task.currentData(); self.task.clear(); self.task.addItem('—',None)
        for r in self.s.rows('SELECT id,name FROM tasks WHERE project_id=? ORDER BY name',(self.proj.currentData(),)): self.task.addItem(r['name'],r['id'])
        i=self.task.findData(old); self.task.setCurrentIndex(max(i,0))
    def data(self):
        d=self.day.date(); a=self.st.time(); z=self.en.time(); st=datetime(d.year(),d.month(),d.day(),a.hour(),a.minute(),a.second()); en=datetime(d.year(),d.month(),d.day(),z.hour(),z.minute(),z.second());
        if en<=st: en+=timedelta(days=1)
        return (self.desc.text().strip(),self.proj.currentData(),self.task.currentData(),st.isoformat(timespec='seconds'),en.isoformat(timespec='seconds'),int((en-st).total_seconds()),self.tags.text().strip(),int(self.bill.isChecked()))

class Main(QMainWindow):
    def __init__(self):
        super().__init__(); self.s=Store(); self.running=None; self.run_started=None; self.auto_last=None; self.auto_started=None; self.idle_alerted=False; self.setWindowTitle(APP); self.resize(1280,800); self.setMinimumSize(1050,680)
        self.build(); self.apply_theme(); self.restore_running(); self.tick=QTimer(self); self.tick.timeout.connect(self.every_second); self.tick.start(1000); self.auto=QTimer(self); self.auto.timeout.connect(self.auto_sample); self.auto.start(5000); self.backup_timer=QTimer(self); self.backup_timer.timeout.connect(self.check_auto_backup); self.backup_timer.start(60000); QTimer.singleShot(5000,self.check_auto_backup)
        self.tray=QSystemTrayIcon(self.style().standardIcon(QStyle.SP_ComputerIcon),self); m=QMenu(); a=m.addAction('Show'); a.triggered.connect(self.showNormal); t=m.addAction('Start / stop timer'); t.triggered.connect(self.toggle); q=m.addAction('Quit'); q.triggered.connect(QApplication.quit); self.tray.setContextMenu(m); self.tray.show()
        self.shortcut=QShortcut(QKeySequence('Ctrl+Alt+T'),self); self.shortcut.activated.connect(self.toggle)
        if keyboard and os.name=='nt':
            try: keyboard.add_hotkey('ctrl+alt+t',lambda: QTimer.singleShot(0,self.toggle))
            except Exception: pass
    def build(self):
        root=QWidget(); self.setCentralWidget(root); h=QHBoxLayout(root); h.setContentsMargins(0,0,0,0); h.setSpacing(0)
        side=QFrame(); side.setObjectName('side'); side.setFixedWidth(210); sv=QVBoxLayout(side); logo=QLabel('◷  ChronoFlow'); logo.setObjectName('logo'); sv.addWidget(logo)
        self.stack=QStackedWidget(); self.nav=[]
        pages=[('Timer',self.timer_page),('Timesheet',self.timesheet_page),('Calendar',self.calendar_page),('Auto Tracker',self.auto_page),('Dashboard',self.dashboard_page),('Reports',self.reports_page),('Projects',self.projects_page),('Settings',self.settings_page)]
        for i,(name,fn) in enumerate(pages):
            b=QPushButton(name); b.setCheckable(True); b.clicked.connect(lambda _,x=i:self.goto(x)); sv.addWidget(b); self.nav.append(b); self.stack.addWidget(fn())
        sv.addStretch(); h.addWidget(side); h.addWidget(self.stack,1); self.goto(0)
    def goto(self,i):
        self.stack.setCurrentIndex(i); [b.setChecked(k==i) for k,b in enumerate(self.nav)]; [self.refresh_timer,self.refresh_timesheet,self.refresh_calendar,self.refresh_auto,self.refresh_dashboard,self.refresh_reports,self.refresh_projects,self.refresh_settings][i]()
    def title(self,text,sub=''):
        box=QVBoxLayout(); a=QLabel(text); a.setObjectName('pageTitle'); box.addWidget(a); 
        if sub: b=QLabel(sub); b.setObjectName('muted'); box.addWidget(b)
        return box
    def timer_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Timer','Track focused work, breaks, and manual entries.'))
        card=QFrame(); card.setObjectName('card'); g=QGridLayout(card); self.desc=QLineEdit(); self.desc.setPlaceholderText('What are you working on?'); self.pcombo=QComboBox(); self.tcombo=QComboBox(); self.tagline=QLineEdit(); self.tagline.setPlaceholderText('tags'); self.bigtime=QLabel('00:00:00'); self.bigtime.setObjectName('bigtime'); self.start=QPushButton('▶  Start'); self.start.setObjectName('primary'); self.start.clicked.connect(self.toggle); self.breakbtn=QPushButton('☕ Break'); self.breakbtn.clicked.connect(self.start_break); self.manual=QPushButton('+ Manual entry'); self.manual.clicked.connect(self.add_manual)
        g.addWidget(self.desc,0,0,1,3); g.addWidget(self.pcombo,1,0); g.addWidget(self.tcombo,1,1); g.addWidget(self.tagline,1,2); g.addWidget(self.bigtime,2,0); g.addWidget(self.start,2,1); g.addWidget(self.breakbtn,2,2); g.addWidget(self.manual,3,2); v.addWidget(card)
        self.pcombo.currentIndexChanged.connect(self.load_tasks)
        self.entries=QTableWidget(0,7); self.entries.setHorizontalHeaderLabels(['Description','Project','Start','End','Duration','Tags','']); self.entries.horizontalHeader().setSectionResizeMode(0,QHeaderView.Stretch); self.entries.setSelectionBehavior(QTableWidget.SelectRows); self.entries.cellDoubleClicked.connect(self.edit_entry); v.addWidget(self.entries,1); return w
    def timesheet_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Timesheet','Weekly totals by project and day.')); self.weeklabel=QLabel(); self.ts=QTableWidget(); v.addWidget(self.weeklabel); v.addWidget(self.ts,1); b=QPushButton('+ Add time'); b.clicked.connect(self.add_manual); v.addWidget(b); return w
    def calendar_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Calendar','Day/week timeline of tracked time.')); bar=QHBoxLayout(); self.calview=QComboBox(); self.calview.addItems(['Week','Day']); self.calview.currentIndexChanged.connect(self.refresh_calendar); bar.addWidget(self.calview); bar.addStretch(); b=QPushButton('+ Block time'); b.clicked.connect(self.add_manual); bar.addWidget(b); v.addLayout(bar); self.cal=QTableWidget(); v.addWidget(self.cal,1); return w
    def auto_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Auto Tracker','Private local activity history from the active Windows app.')); bar=QHBoxLayout(); self.autotoggle=QCheckBox('Record activity automatically'); self.autotoggle.stateChanged.connect(lambda x:self.s.set('auto',1 if x else 0)); self.autosearch=QLineEdit(); self.autosearch.setPlaceholderText('Search app or window title'); self.autosearch.textChanged.connect(self.refresh_auto); bar.addWidget(self.autotoggle); bar.addWidget(self.autosearch,1); v.addLayout(bar); self.autotable=QTableWidget(0,7); self.autotable.setHorizontalHeaderLabels(['','App','Window / document','Start','End','Duration','Idle']); self.autotable.horizontalHeader().setSectionResizeMode(2,QHeaderView.Stretch); v.addWidget(self.autotable,1); bb=QHBoxLayout(); add=QPushButton('Convert selected to time entries'); add.clicked.connect(self.convert_activity); delete=QPushButton('Delete selected'); delete.clicked.connect(self.delete_activity); bb.addWidget(add); bb.addWidget(delete); bb.addStretch(); v.addLayout(bb); return w
    def dashboard_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Dashboard','Where your time is going.')); self.cards=QHBoxLayout(); v.addLayout(self.cards); charts=QHBoxLayout(); self.pie=QChartView(); self.bar=QChartView(); charts.addWidget(self.pie); charts.addWidget(self.bar); v.addLayout(charts,1); return w
    def reports_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Reports','Filter, analyze, and export your history.')); bar=QHBoxLayout(); self.rf=QDateEdit(QDate.currentDate().addDays(-30)); self.rt=QDateEdit(QDate.currentDate()); self.rp=QComboBox(); self.rsearch=QLineEdit(); self.rsearch.setPlaceholderText('Description/tag'); run=QPushButton('Apply'); run.clicked.connect(self.refresh_reports); [bar.addWidget(x) for x in [self.rf,self.rt,self.rp,self.rsearch,run]]; v.addLayout(bar); self.rsummary=QLabel(); self.rtable=QTableWidget(); v.addWidget(self.rsummary); v.addWidget(self.rtable,1); ex=QHBoxLayout();
        for n,fn in [('Export CSV',self.export_csv),('Export Excel',self.export_xlsx),('Export PDF',self.export_pdf)]: b=QPushButton(n); b.clicked.connect(fn); ex.addWidget(b)
        ex.addStretch(); v.addLayout(ex); return w
    def projects_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Projects','Projects, tasks, rates, estimates, and archive.')); self.ptable=QTableWidget(); v.addWidget(self.ptable,1); b=QHBoxLayout();
        for n,fn in [('+ Project',self.add_project),('+ Task',self.add_task),('Archive / restore',self.archive_project)]: x=QPushButton(n); x.clicked.connect(fn); b.addWidget(x)
        b.addStretch(); v.addLayout(b); return w
    def settings_page(self):
        w=QWidget(); v=QVBoxLayout(w); v.setContentsMargins(30,25,30,25); v.addLayout(self.title('Settings','Local-first preferences.')); f=QFormLayout(); self.theme=QComboBox(); self.theme.addItems(['Dark','Light']); self.idle=QSpinBox(); self.idle.setRange(1,120); self.remind=QSpinBox(); self.remind.setRange(0,240); self.pomo=QSpinBox(); self.pomo.setRange(1,120); self.pbreak=QSpinBox(); self.pbreak.setRange(1,60); self.defproj=QComboBox(); self.startup=QCheckBox('Launch ChronoFlow at login'); self.excluded=QLineEdit(); self.excluded.setPlaceholderText('e.g. LockApp.exe, SearchHost.exe');
        for n,x in [('Theme',self.theme),('Idle threshold (minutes)',self.idle),('Reminder when timer is off (minutes, 0=off)',self.remind),('Pomodoro work (minutes)',self.pomo),('Pomodoro break (minutes)',self.pbreak),('Default project',self.defproj),('Auto Tracker excluded apps',self.excluded),('',self.startup)]: f.addRow(n,x)
        save=QPushButton('Save settings'); save.setObjectName('primary'); save.clicked.connect(self.save_settings); f.addRow(save); v.addLayout(f)
        box=QFrame(); box.setObjectName('card'); bf=QFormLayout(box); title=QLabel('Automatic backup'); title.setStyleSheet('font-size:16px;font-weight:700'); bf.addRow(title)
        self.backup_auto=QCheckBox('Enable automatic backups'); self.backup_schedule=QComboBox(); self.backup_schedule.addItems(['Every 1 hour','Every 6 hours','Every 12 hours','Daily','Every 3 days','Weekly','Custom hours','Custom days','When app exits']); self.backup_custom=QSpinBox(); self.backup_custom.setRange(1,365); self.backup_custom.setSuffix(' hours'); self.backup_schedule.currentTextChanged.connect(self.update_backup_custom_ui)
        self.backup_path=QLineEdit(); choose=QPushButton('Browse…'); choose.clicked.connect(self.choose_backup_folder); ph=QHBoxLayout(); ph.addWidget(self.backup_path,1); ph.addWidget(choose); pw=QWidget(); pw.setLayout(ph)
        self.backup_keep=QSpinBox(); self.backup_keep.setRange(1,1000); self.backup_keep.setSuffix(' backups'); self.backup_last=QLabel('Never')
        bf.addRow('',self.backup_auto); bf.addRow('Backup schedule',self.backup_schedule); bf.addRow('Custom interval',self.backup_custom); bf.addRow('Backup folder',pw); bf.addRow('Keep newest',self.backup_keep); bf.addRow('Last successful backup',self.backup_last)
        row=QHBoxLayout(); bk=QPushButton('Back up now'); bk.clicked.connect(self.backup_now); rs=QPushButton('Restore from backup'); rs.clicked.connect(self.restore); op=QPushButton('Open backup folder'); op.clicked.connect(self.open_backup_folder); row.addWidget(bk); row.addWidget(rs); row.addWidget(op); row.addStretch(); rw=QWidget(); rw.setLayout(row); bf.addRow(rw); v.addWidget(box); v.addStretch(); return w
    def apply_theme(self):
        dark=self.s.get('theme','Dark')=='Dark'; bg='#0f1117' if dark else '#f5f7fb'; panel='#171a23' if dark else '#ffffff'; text='#f2f4f8' if dark else '#172033'; muted='#9198aa' if dark else '#667085'; border='#2a2f3a' if dark else '#dfe3ea'; inp='#11141b' if dark else '#f8fafc'
        qss = """
QWidget { background: BG; color: TEXT; font-family: "Segoe UI"; font-size: 13px; }
#side { background: PANEL; border-right: 1px solid BORDER; }
#logo { font-size: 20px; font-weight: 700; padding: 18px 8px; }
#pageTitle { font-size: 26px; font-weight: 700; }
#muted { color: MUTED; margin-bottom: 10px; }
#card { background: PANEL; border: 1px solid BORDER; border-radius: 12px; padding: 12px; }
QLineEdit,QComboBox,QDateEdit,QTimeEdit,QSpinBox,QDoubleSpinBox { background: INP; border: 1px solid BORDER; border-radius: 7px; padding: 8px; }
QPushButton { background: PANEL; border: 1px solid BORDER; border-radius: 7px; padding: 9px 13px; text-align: left; }
QPushButton:hover { border-color: #5b8def; }
QPushButton:checked,#primary { background: #5b8def; color: white; border-color: #5b8def; }
#bigtime { font-size: 28px; font-weight: 700; }
QTableWidget { background: PANEL; border: 1px solid BORDER; border-radius: 9px; gridline-color: BORDER; }
QHeaderView::section { background: PANEL; padding: 8px; border: 0; border-bottom: 1px solid BORDER; font-weight: 600; }
"""
        qss = qss.replace('BG',bg).replace('TEXT',text).replace('PANEL',panel).replace('BORDER',border).replace('MUTED',muted).replace('INP',inp)
        self.setStyleSheet(qss)

    def load_projects(self):
        cur=self.pcombo.currentData(); self.pcombo.blockSignals(True); self.pcombo.clear();
        for r in self.s.rows('SELECT id,name FROM projects WHERE archived=0 ORDER BY name'): self.pcombo.addItem(r['name'],r['id'])
        i=self.pcombo.findData(cur or int(self.s.get('default_project','0') or 0)); self.pcombo.setCurrentIndex(max(i,0)); self.pcombo.blockSignals(False); self.load_tasks()
    def load_tasks(self):
        self.tcombo.clear(); self.tcombo.addItem('No task',None)
        for r in self.s.rows('SELECT id,name FROM tasks WHERE project_id=? ORDER BY name',(self.pcombo.currentData(),)): self.tcombo.addItem(r['name'],r['id'])
    def toggle(self):
        if self.running: self.stop_timer()
        else: self.start_timer()
    def start_timer(self,is_break=False):
        if self.running: return
        desc='Break' if is_break else self.desc.text().strip(); pid=self.pcombo.currentData(); tid=None if is_break else self.tcombo.currentData(); tags='' if is_break else self.tagline.text().strip(); st=nowiso(); self.running=self.s.exec('INSERT INTO entries(description,project_id,task_id,start,tags,is_break) VALUES(?,?,?,?,?,?)',(desc,pid,tid,st,tags,int(is_break))); self.run_started=parse_dt(st); self.s.set('running',self.running); self.start.setText('■  Stop'); self.refresh_timer()
    def start_break(self):
        if self.running: self.stop_timer()
        self.start_timer(True)
    def stop_timer(self):
        if not self.running:return
        en=datetime.now(); sec=int((en-self.run_started).total_seconds()); self.s.exec('UPDATE entries SET end=?,duration=? WHERE id=?',(en.isoformat(timespec='seconds'),sec,self.running)); self.running=None; self.run_started=None; self.s.set('running',''); self.start.setText('▶  Start'); self.bigtime.setText('00:00:00'); self.refresh_timer()
    def restore_running(self):
        x=self.s.get('running','');
        if x:
            r=self.s.one('SELECT * FROM entries WHERE id=? AND end IS NULL',(x,));
            if r: self.running=r['id']; self.run_started=parse_dt(r['start']); self.start.setText('■  Stop')
    def every_second(self):
        if self.running:
            self.bigtime.setText(durfmt((datetime.now()-self.run_started).total_seconds())); threshold=int(self.s.get('idle','10'))*60
            if idle_seconds()>threshold and not self.idle_alerted:
                self.idle_alerted=True; self.tray.showMessage('Idle time detected',f'You have been inactive for {int(idle_seconds()/60)} minutes. Open ChronoFlow to review.',QSystemTrayIcon.Information,8000)
            if idle_seconds()<5: self.idle_alerted=False
            pm=int(self.s.get('pomo','25'))*60
            if pm and int((datetime.now()-self.run_started).total_seconds())==pm: self.tray.showMessage('Pomodoro complete',f'Time for a {self.s.get("pbreak","5")} minute break.',QSystemTrayIcon.Information,8000)
        else:
            rem=int(self.s.get('remind','0') or 0)
            if rem and int(time.time())%(rem*60)==0: self.tray.showMessage('Timer is off','Remember to track what you are working on.',QSystemTrayIcon.Information,6000)
    def refresh_timer(self):
        self.load_projects(); rs=self.s.rows('''SELECT e.*,p.name project FROM entries e LEFT JOIN projects p ON p.id=e.project_id ORDER BY e.start DESC LIMIT 100'''); self.entries.setRowCount(len(rs))
        for i,r in enumerate(rs):
            vals=[r['description'] or ('Break' if r['is_break'] else ''),r['project'] or '',parse_dt(r['start']).strftime('%m/%d %H:%M'),parse_dt(r['end']).strftime('%H:%M') if r['end'] else 'Running',durfmt(r['duration'] if r['end'] else (datetime.now()-parse_dt(r['start'])).total_seconds()),r['tags'] or '','▶  ⋯']
            for j,x in enumerate(vals): self.entries.setItem(i,j,QTableWidgetItem(str(x)))
            self.entries.item(i,0).setData(Qt.UserRole,r['id'])
    def add_manual(self):
        d=EntryDialog(self.s,self); 
        if d.exec():
            x=d.data(); self.s.exec('INSERT INTO entries(description,project_id,task_id,start,end,duration,tags,billable) VALUES(?,?,?,?,?,?,?,?)',x); self.goto(self.stack.currentIndex())
    def edit_entry(self,row,col):
        eid=self.entries.item(row,0).data(Qt.UserRole); r=self.s.one('SELECT * FROM entries WHERE id=?',(eid,));
        if not r:return
        d=EntryDialog(self.s,self,r)
        if d.exec():
            x=d.data(); self.s.exec('UPDATE entries SET description=?,project_id=?,task_id=?,start=?,end=?,duration=?,tags=?,billable=? WHERE id=?',x+(eid,)); self.refresh_timer()
    def refresh_timesheet(self):
        today=date.today(); monday=today-timedelta(days=today.weekday()); days=[monday+timedelta(days=i) for i in range(7)]; self.weeklabel.setText(f'{days[0]:%b %d} – {days[-1]:%b %d, %Y}')
        projs=self.s.rows('SELECT id,name FROM projects WHERE archived=0 ORDER BY name'); self.ts.setColumnCount(9); self.ts.setHorizontalHeaderLabels(['Project']+[d.strftime('%a %d') for d in days]+['Total']); self.ts.setRowCount(len(projs)+1)
        totals=[0]*7
        for i,p in enumerate(projs):
            self.ts.setItem(i,0,QTableWidgetItem(p['name'])); rt=0
            for j,d in enumerate(days):
                n=self.s.one("SELECT COALESCE(SUM(duration),0) s FROM entries WHERE project_id=? AND date(start)=?",(p['id'],d.isoformat()))['s']; totals[j]+=n; rt+=n; self.ts.setItem(i,j+1,QTableWidgetItem(durfmt(n)[:-3]))
            self.ts.setItem(i,8,QTableWidgetItem(durfmt(rt)[:-3]))
        self.ts.setItem(len(projs),0,QTableWidgetItem('TOTAL'))
        for j,n in enumerate(totals): self.ts.setItem(len(projs),j+1,QTableWidgetItem(durfmt(n)[:-3]))
        self.ts.setItem(len(projs),8,QTableWidgetItem(durfmt(sum(totals))[:-3])); self.ts.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    def refresh_calendar(self):
        if not hasattr(self,'cal'):return
        today=date.today(); daymode=self.calview.currentText()=='Day'; days=[today] if daymode else [today-timedelta(days=today.weekday())+timedelta(days=i) for i in range(7)]; self.cal.setColumnCount(len(days)); self.cal.setHorizontalHeaderLabels([d.strftime('%a\n%b %d') for d in days]); self.cal.setRowCount(24); self.cal.setVerticalHeaderLabels([f'{h:02d}:00' for h in range(24)])
        for c,d in enumerate(days):
            for r in self.s.rows('SELECT e.*,p.name project FROM entries e LEFT JOIN projects p ON p.id=e.project_id WHERE date(e.start)=? ORDER BY e.start',(d.isoformat(),)):
                st=parse_dt(r['start']); en=parse_dt(r['end']) or datetime.now(); rr=st.hour; txt=f'{st:%H:%M}–{en:%H:%M}\n{r["description"] or "(no description)"}\n{r["project"] or ""}'; old=self.cal.item(rr,c); self.cal.setItem(rr,c,QTableWidgetItem((old.text()+'\n\n' if old else '')+txt))
        self.cal.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); self.cal.verticalHeader().setDefaultSectionSize(46)
    def auto_sample(self):
        if self.s.get('auto','1')!='1': return
        app,title=fg_info(); excluded=[x.strip().lower() for x in self.s.get('excluded','').split(',') if x.strip()]
        if app.lower() in excluded:return
        idle=1 if idle_seconds()>60 else 0; key=(app,title,idle); now=datetime.now()
        if self.auto_last==key and self.auto_started:
            r=self.s.one('SELECT id,start FROM activity ORDER BY id DESC LIMIT 1');
            if r:self.s.exec('UPDATE activity SET end=?,duration=? WHERE id=?',(now.isoformat(timespec='seconds'),int((now-parse_dt(r['start'])).total_seconds()),r['id']))
        else:
            self.s.exec('INSERT INTO activity(app,title,start,end,duration,idle) VALUES(?,?,?,?,?,?)',(app,title,now.isoformat(timespec='seconds'),now.isoformat(timespec='seconds'),0,idle)); self.auto_last=key; self.auto_started=now
    def refresh_auto(self):
        if not hasattr(self,'autotable'):return
        self.autotoggle.blockSignals(True); self.autotoggle.setChecked(self.s.get('auto','1')=='1'); self.autotoggle.blockSignals(False); q=self.autosearch.text().strip(); rs=self.s.rows('SELECT * FROM activity WHERE (app LIKE ? OR title LIKE ?) ORDER BY start DESC LIMIT 1000',(f'%{q}%',f'%{q}%')); self.autotable.setRowCount(len(rs))
        for i,r in enumerate(rs):
            chk=QTableWidgetItem(); chk.setCheckState(Qt.Unchecked); chk.setData(Qt.UserRole,r['id']); self.autotable.setItem(i,0,chk); vals=[r['app'],r['title'],parse_dt(r['start']).strftime('%m/%d %H:%M:%S'),parse_dt(r['end']).strftime('%H:%M:%S'),durfmt(r['duration']),f'{100 if r["idle"] else 0}%']
            for j,x in enumerate(vals,1):self.autotable.setItem(i,j,QTableWidgetItem(str(x)))
    def selected_activity(self): return [self.autotable.item(i,0).data(Qt.UserRole) for i in range(self.autotable.rowCount()) if self.autotable.item(i,0).checkState()==Qt.Checked]
    def convert_activity(self):
        ids=self.selected_activity();
        for aid in ids:
            r=self.s.one('SELECT * FROM activity WHERE id=?',(aid,)); self.s.exec('INSERT INTO entries(description,project_id,start,end,duration) VALUES(?,?,?,?,?)',(r['title'] or r['app'],self.pcombo.currentData(),r['start'],r['end'],r['duration'])); self.s.exec('UPDATE activity SET converted=1 WHERE id=?',(aid,))
        self.refresh_auto(); self.tray.showMessage(APP,f'Converted {len(ids)} activities to time entries.')
    def delete_activity(self):
        for aid in self.selected_activity(): self.s.exec('DELETE FROM activity WHERE id=?',(aid,))
        self.refresh_auto()
    def clear_layout(self,l):
        while l.count():
            it=l.takeAt(0); w=it.widget();
            if w:w.deleteLater()
    def refresh_dashboard(self):
        self.clear_layout(self.cards); today=date.today(); mon=today-timedelta(days=today.weekday()); month=today.replace(day=1)
        for label,start in [('Today',today),('This week',mon),('This month',month)]:
            sec=self.s.one('SELECT COALESCE(SUM(duration),0)s FROM entries WHERE date(start)>=?',(start.isoformat(),))['s']; c=QFrame(); c.setObjectName('card'); l=QVBoxLayout(c); l.addWidget(QLabel(label)); x=QLabel(durfmt(sec)[:-3]); x.setObjectName('bigtime'); l.addWidget(x); self.cards.addWidget(c)
        ps=QPieSeries();
        for r in self.s.rows('''SELECT COALESCE(p.name,'No project') n,SUM(e.duration)s FROM entries e LEFT JOIN projects p ON p.id=e.project_id WHERE date(e.start)>=? GROUP BY n''',(mon.isoformat(),)): ps.append(r['n'],r['s'])
        ch=QChart(); ch.addSeries(ps); ch.setTitle('This week by project'); ch.legend().setVisible(True); self.pie.setChart(ch)
        bs=QBarSeries(); bset=QBarSet('Hours'); cats=[]
        for i in range(7):
            d=mon+timedelta(days=i); cats.append(d.strftime('%a')); sec=self.s.one('SELECT COALESCE(SUM(duration),0)s FROM entries WHERE date(start)=?',(d.isoformat(),))['s']; bset.append(sec/3600)
        bs.append(bset); bc=QChart(); bc.addSeries(bs); ax=QBarCategoryAxis(); ax.append(cats); bc.addAxis(ax,Qt.AlignBottom); bs.attachAxis(ax); ay=QValueAxis(); ay.setTitleText('Hours'); bc.addAxis(ay,Qt.AlignLeft); bs.attachAxis(ay); bc.setTitle('Daily hours'); self.bar.setChart(bc)
    def report_rows(self):
        a=self.rf.date().toPython().isoformat(); z=self.rt.date().toPython().isoformat(); pid=self.rp.currentData(); q=self.rsearch.text().strip(); sql='''SELECT e.*,p.name project,t.name task FROM entries e LEFT JOIN projects p ON p.id=e.project_id LEFT JOIN tasks t ON t.id=e.task_id WHERE date(e.start) BETWEEN ? AND ?'''; p=[a,z]
        if pid: sql+=' AND e.project_id=?'; p.append(pid)
        if q: sql+=' AND (e.description LIKE ? OR e.tags LIKE ?)'; p += [f'%{q}%',f'%{q}%']
        return self.s.rows(sql+' ORDER BY e.start DESC',p)
    def refresh_reports(self):
        if not hasattr(self,'rtable'):return
        cur=self.rp.currentData(); self.rp.clear(); self.rp.addItem('All projects',None)
        for r in self.s.rows('SELECT id,name FROM projects ORDER BY name'):self.rp.addItem(r['name'],r['id'])
        i=self.rp.findData(cur); self.rp.setCurrentIndex(max(i,0)); rs=self.report_rows(); total=sum(r['duration'] for r in rs); amount=sum((r['duration']/3600)*(self.s.one('SELECT rate FROM projects WHERE id=?',(r['project_id'],))['rate'] if r['project_id'] else 0) for r in rs); self.rsummary.setText(f'{len(rs)} entries  •  {durfmt(total)[:-3]}  •  Estimated value {amount:,.2f}')
        self.rtable.setColumnCount(8); self.rtable.setHorizontalHeaderLabels(['Date','Description','Project','Task','Tags','Start','End','Duration']); self.rtable.setRowCount(len(rs))
        for i,r in enumerate(rs):
            st=parse_dt(r['start']); en=parse_dt(r['end']); vals=[st.strftime('%Y-%m-%d'),r['description'] or '',r['project'] or '',r['task'] or '',r['tags'] or '',st.strftime('%H:%M'),en.strftime('%H:%M') if en else '',durfmt(r['duration'])]
            for j,x in enumerate(vals):self.rtable.setItem(i,j,QTableWidgetItem(str(x)))
        self.rtable.horizontalHeader().setSectionResizeMode(1,QHeaderView.Stretch)
    def export_csv(self):
        p,_=QFileDialog.getSaveFileName(self,'Export CSV','ChronoFlow-report.csv','CSV (*.csv)');
        if not p:return
        with open(p,'w',newline='',encoding='utf-8-sig') as f:
            w=csv.writer(f); w.writerow(['Date','Description','Project','Task','Tags','Start','End','Duration seconds']);
            for r in self.report_rows():w.writerow([r['start'][:10],r['description'],r['project'],r['task'],r['tags'],r['start'],r['end'],r['duration']])
    def export_xlsx(self):
        p,_=QFileDialog.getSaveFileName(self,'Export Excel','ChronoFlow-report.xlsx','Excel (*.xlsx)');
        if not p:return
        wb=Workbook(); ws=wb.active; ws.title='Time entries'; ws.append(['Date','Description','Project','Task','Tags','Start','End','Duration (hours)']);
        for r in self.report_rows():ws.append([r['start'][:10],r['description'],r['project'],r['task'],r['tags'],r['start'],r['end'],r['duration']/3600]); wb.save(p)
    def export_pdf(self):
        p,_=QFileDialog.getSaveFileName(self,'Export PDF','ChronoFlow-report.pdf','PDF (*.pdf)');
        if not p:return
        c=canvas.Canvas(p,pagesize=A4); W,H=A4; y=H-50; c.setFont('Helvetica-Bold',18); c.drawString(40,y,'ChronoFlow Time Report'); y-=30; c.setFont('Helvetica',9)
        for r in self.report_rows():
            line=f"{r['start'][:10]}  {r['description'] or ''}  |  {r['project'] or ''}  |  {durfmt(r['duration'])}"
            c.drawString(40,y,line[:110]); y-=14
            if y<50:c.showPage();y=H-50;c.setFont('Helvetica',9)
        c.save()
    def refresh_projects(self):
        rs=self.s.rows('SELECT * FROM projects ORDER BY archived,name'); self.ptable.setColumnCount(6); self.ptable.setHorizontalHeaderLabels(['Project','Color','Hourly rate','Estimate h','Status','Tasks']); self.ptable.setRowCount(len(rs))
        for i,r in enumerate(rs):
            tasks=', '.join(x['name'] for x in self.s.rows('SELECT name FROM tasks WHERE project_id=?',(r['id'],))); vals=[r['name'],r['color'],f'{r["rate"]:.2f}',f'{r["estimate"]:.1f}','Archived' if r['archived'] else 'Active',tasks]
            for j,x in enumerate(vals):self.ptable.setItem(i,j,QTableWidgetItem(str(x))); self.ptable.item(i,0).setData(Qt.UserRole,r['id'])
        self.ptable.horizontalHeader().setSectionResizeMode(5,QHeaderView.Stretch)
    def add_project(self):
        d=QDialog(self); d.setWindowTitle('New project'); f=QFormLayout(d); name=QLineEdit(); color=QLineEdit('#5B8DEF'); rate=QDoubleSpinBox(); rate.setMaximum(1e7); est=QDoubleSpinBox(); est.setMaximum(1e6); b=QPushButton('Create'); b.clicked.connect(d.accept); [f.addRow(n,w) for n,w in [('Name',name),('Color',color),('Hourly rate',rate),('Estimate hours',est),('',b)]]
        if d.exec() and name.text().strip():
            try:self.s.exec('INSERT INTO projects(name,color,rate,estimate) VALUES(?,?,?,?)',(name.text().strip(),color.text(),rate.value(),est.value()))
            except sqlite3.IntegrityError:pass
            self.refresh_projects()
    def selected_project_id(self):
        r=self.ptable.currentRow(); return self.ptable.item(r,0).data(Qt.UserRole) if r>=0 else None
    def add_task(self):
        pid=self.selected_project_id();
        if not pid:return
        n,ok=QInputDialog.getText(self,'New task','Task name');
        if ok and n.strip():
            try:self.s.exec('INSERT INTO tasks(project_id,name) VALUES(?,?)',(pid,n.strip()))
            except sqlite3.IntegrityError:pass
            self.refresh_projects()
    def archive_project(self):
        pid=self.selected_project_id();
        if pid:self.s.exec('UPDATE projects SET archived=1-archived WHERE id=?',(pid,));self.refresh_projects()
    def refresh_settings(self):
        self.theme.setCurrentText(self.s.get('theme','Dark')); self.idle.setValue(int(self.s.get('idle','10'))); self.remind.setValue(int(self.s.get('remind','0'))); self.pomo.setValue(int(self.s.get('pomo','25'))); self.pbreak.setValue(int(self.s.get('pbreak','5'))); self.excluded.setText(self.s.get('excluded','')); self.defproj.clear();
        for r in self.s.rows('SELECT id,name FROM projects WHERE archived=0 ORDER BY name'):self.defproj.addItem(r['name'],r['id'])
        self.defproj.setCurrentIndex(max(self.defproj.findData(int(self.s.get('default_project','0') or 0)),0)); self.startup.setChecked(self.startup_enabled())
        self.backup_auto.setChecked(self.s.get('backup_auto','1')=='1'); self.backup_schedule.setCurrentText(self.s.get('backup_schedule','Daily')); self.backup_custom.setValue(int(self.s.get('backup_custom','24'))); self.backup_path.setText(self.s.get('backup_path',str(DEFAULT_BACKUP_DIR))); self.backup_keep.setValue(int(self.s.get('backup_keep','30'))); self.backup_last.setText(self.s.get('backup_last','Never')); self.update_backup_custom_ui()
    def save_settings(self):
        for k,v in [('theme',self.theme.currentText()),('idle',self.idle.value()),('remind',self.remind.value()),('pomo',self.pomo.value()),('pbreak',self.pbreak.value()),('default_project',self.defproj.currentData()),('excluded',self.excluded.text()),('backup_auto',1 if self.backup_auto.isChecked() else 0),('backup_schedule',self.backup_schedule.currentText()),('backup_custom',self.backup_custom.value()),('backup_path',self.backup_path.text().strip() or str(DEFAULT_BACKUP_DIR)),('backup_keep',self.backup_keep.value())]:self.s.set(k,v)
        self.set_startup(self.startup.isChecked()); self.apply_theme(); QMessageBox.information(self,'Saved','Settings saved.')
    def startup_enabled(self):
        if sys.platform=='darwin':return (Path.home()/'Library'/'LaunchAgents'/'com.chronoflow.login.plist').exists()
        if os.name!='nt':return False
        try:
            import winreg; k=winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Run'); winreg.QueryValueEx(k,APP); return True
        except Exception:return False
    def set_startup(self,on):
        if sys.platform=='darwin':
            import plistlib
            folder=Path.home()/'Library'/'LaunchAgents'; target=folder/'com.chronoflow.login.plist'
            if on:
                folder.mkdir(parents=True,exist_ok=True)
                exe=Path(sys.executable).resolve()
                bundle=next((parent for parent in exe.parents if parent.suffix=='.app'),None)
                args=['/usr/bin/open','-a',str(bundle)] if bundle else [str(exe),str(Path(__file__).resolve())]
                target.write_bytes(plistlib.dumps({'Label':'com.chronoflow.login','ProgramArguments':args,'RunAtLoad':True}))
            else:target.unlink(missing_ok=True)
            return
        if os.name!='nt':return
        try:
            import winreg; k=winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Run',0,winreg.KEY_SET_VALUE)
            if on:winreg.SetValueEx(k,APP,0,winreg.REG_SZ,f'"{sys.executable}"')
            else:
                try:winreg.DeleteValue(k,APP)
                except FileNotFoundError:pass
        except Exception:pass
    def update_backup_custom_ui(self):
        mode=self.backup_schedule.currentText(); custom=mode in ('Custom hours','Custom days'); self.backup_custom.setEnabled(custom); self.backup_custom.setSuffix(' hours' if mode=='Custom hours' else ' days')
    def choose_backup_folder(self):
        p=QFileDialog.getExistingDirectory(self,'Choose backup folder',self.backup_path.text() or str(DEFAULT_BACKUP_DIR))
        if p:self.backup_path.setText(p)
    def open_backup_folder(self):
        folder=Path(self.backup_path.text().strip() or self.s.get('backup_path',str(DEFAULT_BACKUP_DIR))); folder.mkdir(parents=True,exist_ok=True)
        try:
            if os.name=='nt': os.startfile(str(folder))
            elif sys.platform=='darwin': os.system(f'open "{folder}"')
            else: os.system(f'xdg-open "{folder}" >/dev/null 2>&1 &')
        except Exception: pass
    def backup_interval_seconds(self):
        mode=self.s.get('backup_schedule','Daily'); custom=max(1,int(self.s.get('backup_custom','24'))); values={'Every 1 hour':3600,'Every 6 hours':21600,'Every 12 hours':43200,'Daily':86400,'Every 3 days':259200,'Weekly':604800,'Custom hours':custom*3600,'Custom days':custom*86400}; return values.get(mode)
    def create_backup(self,folder=None,notify=False):
        folder=Path(folder or self.s.get('backup_path',str(DEFAULT_BACKUP_DIR))); folder.mkdir(parents=True,exist_ok=True); stamp=datetime.now().strftime('%Y-%m-%d_%H-%M-%S'); target=folder/f'ChronoFlow_{stamp}.db'
        try:
            self.s.c.commit(); dst=sqlite3.connect(target); self.s.c.backup(dst); dst.close(); self.s.set('backup_last',nowiso()); self.cleanup_backups(folder)
            if hasattr(self,'backup_last'): self.backup_last.setText(self.s.get('backup_last'))
            if notify: QMessageBox.information(self,'Backup complete',f'Backup saved to:\n{target}')
            return target
        except Exception as ex:
            if notify: QMessageBox.critical(self,'Backup failed',str(ex))
            return None
    def cleanup_backups(self,folder):
        keep=max(1,int(self.s.get('backup_keep','30'))); files=sorted(Path(folder).glob('ChronoFlow_*.db'),key=lambda x:x.stat().st_mtime,reverse=True)
        for p in files[keep:]:
            try:p.unlink()
            except Exception:pass
    def check_auto_backup(self):
        if self.s.get('backup_auto','1')!='1':return
        mode=self.s.get('backup_schedule','Daily'); interval=self.backup_interval_seconds()
        if mode=='When app exits' or not interval:return
        last=self.s.get('backup_last',''); due=True
        if last:
            try:due=(datetime.now()-parse_dt(last)).total_seconds()>=interval
            except Exception:due=True
        if due:self.create_backup()
    def backup_now(self):
        folder=self.backup_path.text().strip() if hasattr(self,'backup_path') else self.s.get('backup_path',str(DEFAULT_BACKUP_DIR)); self.create_backup(folder,True)
    def backup(self): self.backup_now()
    def restore(self):
        p,_=QFileDialog.getOpenFileName(self,'Restore database','','Database (*.db)');
        if p and QMessageBox.question(self,'Restore','Replace all current data with this backup?')==QMessageBox.Yes:self.s.c.close();shutil.copy2(p,DB);QMessageBox.information(self,'Restart required','Restart ChronoFlow to load the restored database.')
    def closeEvent(self,e):
        if self.s.get('close_to_tray','1')=='1': e.ignore(); self.hide(); self.tray.showMessage(APP,'ChronoFlow is still running in the system tray.',QSystemTrayIcon.Information,2500)
        else:
            if self.s.get('backup_auto','1')=='1' and self.s.get('backup_schedule','Daily')=='When app exits': self.create_backup()
            e.accept()

if __name__=='__main__':
    app=QApplication(sys.argv); app.setApplicationName(APP); app.setQuitOnLastWindowClosed(False); w=Main(); w.show(); sys.exit(app.exec())
