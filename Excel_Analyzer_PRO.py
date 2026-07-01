# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════╗
║   Excel Analyzer PRO  v5.0  —  Expert All-in-One Edition    ║
║   Features:                                                  ║
║   * Single/Multi file analysis                               ║
║   * Secured Excel direct load (Windows COM / fallback)       ║
║   * Contract-number based merge                              ║
║   * Multi-column append from external file                   ║
║   * Branch name replacement / normalization                  ║
║   * Duplicate & blank cell inspection                        ║
║   * KPI cards, Charts, Search, Filter                        ║
║   * Result Excel download                                    ║
║   * Mac / Windows compatible                                 ║
╚══════════════════════════════════════════════════════════════╝
Run  : python Excel_Analyzer_PRO.py
Needs: pip install pandas openpyxl xlrd matplotlib
"""
import sys, os, threading, subprocess, json, platform, re
from datetime import datetime

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LNX = not IS_WIN and not IS_MAC

# ── auto-install (macOS --break-system-packages 자동 처리) ──────
def _install(pkg):
    for cmd in [
        [sys.executable,"-m","pip","install",pkg,"-q"],
        [sys.executable,"-m","pip","install",pkg,"-q","--break-system-packages"],
        [sys.executable,"-m","pip","install",pkg,"-q","--user"],
    ]:
        try:
            subprocess.check_call(cmd,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            return True
        except subprocess.CalledProcessError:
            continue
    return False

_PKGS = ["pandas","openpyxl","xlrd","matplotlib","requests","urllib3"]
if IS_WIN:  _PKGS += ["xlwings","pywin32"]
elif IS_MAC: _PKGS += ["xlwings"]
for _p in _PKGS:
    try: __import__(_p.split("[")[0])
    except ImportError:
        print(f"[설치중] {_p}..."); _install(_p)

# ── matplotlib backend: use() 반드시 import 전에 호출 ───────────
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt, matplotlib.font_manager as fm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import warnings; warnings.filterwarnings("ignore")

# ── Mac 시스템 폰트 (Apple SD Gothic Neo: 한글 선명) ────────────
FN = ("Malgun Gothic" if IS_WIN
      else "Apple SD Gothic Neo" if IS_MAC
      else "NanumGothic")

# ── matplotlib 한글 폰트 ─────────────────────────────────────────
def _set_font():
    candidates = (["Malgun Gothic","Arial Unicode MS"] if IS_WIN
                  else ["Apple SD Gothic Neo","AppleGothic","Arial Unicode MS"] if IS_MAC
                  else ["NanumGothic","DejaVu Sans"])
    for n in candidates:
        if any(n.lower() in f.name.lower() for f in fm.fontManager.ttflist):
            plt.rcParams["font.family"] = n
            plt.rcParams["axes.unicode_minus"] = False
            return n
_set_font()

# ── Palette ──────────────────────────────────────────────────────
C = dict(
    bg="#0d1117",surface="#161b22",card="#1c2333",card2="#21262d",
    border="#30363d",accent="#58a6ff",accent_h="#79c0ff",
    green="#3fb950",green_h="#56d364",purple="#bc8cff",purple_h="#d2a8ff",
    orange="#ffa657",orange_h="#ffb77c",red="#f85149",red_h="#ff7b72",
    cyan="#39d353",yellow="#e3b341",teal="#56d4c8",
    text="#e6edf3",text2="#c9d1d9",muted="#7d8590",muted2="#484f58",
)
CHART=["#58a6ff","#3fb950","#bc8cff","#ffa657","#f85149",
       "#39d353","#e3b341","#79c0ff","#56d364","#ff7b72",
       "#ffb77c","#d2a8ff","#a5d6ff","#ffa198","#b3f0ff"]

OPS=[("포함","ct"),("미포함","nct"),("같음","eq"),("다름","ne"),
     (">","gt"),("<","lt"),(">=","gte"),("<=","lte"),
     ("시작","sw"),("끝","ew"),("비어있음","emp"),("비어있지않음","nemp"),("정규식","rx")]
AGG={"합계":"sum","평균":"mean","건수":"count","최대":"max","최소":"min","중앙값":"median"}

RECENT=os.path.join(os.path.expanduser("~"),".excel_pro5_recent.json")
def _load_recent():
    try:
        with open(RECENT) as f: return json.load(f)
    except: return []
def _save_recent(p,lst):
    lst=[p]+[x for x in lst if x!=p and os.path.exists(x)]
    try:
        with open(RECENT,"w") as f: json.dump(lst[:8],f)
    except: pass

# ============================================================
# ENGINE 1: Virtual Scroll Tree  (O(1) render regardless of rows)
# ============================================================
class VTree(ttk.Treeview):
    PAGE=500
    def __init__(self,master,**kw):
        super().__init__(master,**kw)
        self._vd=[];self._vc=[];self._vo=0
        self._sc=None;self._sa=True
        self.bind("<Button-1>",self._hclick)
        self.tag_configure("odd", background="#131820")
        self.tag_configure("even",background=C["card"])
        self.tag_configure("dup", background="#3d1515",foreground=C["red_h"])
        self.tag_configure("blank",background="#2d2a00",foreground=C["yellow"])

    def load(self,df,cols):
        self._vc=[c for c in cols if c in df.columns]
        self["columns"]=self._vc
        for c in self._vc:
            lbl=c+(" ^" if c==self._sc and self._sa else " v" if c==self._sc else "")
            self.heading(c,text=lbl)
            self.column(c,width=max(70,min(220,max(len(c)*9,80))),minwidth=40,anchor="w")
        # 컬럼별 ID 여부 미리 계산 (계약번호/서비스번호/청약번호 등 콤마 제거)
        _nc=[_is_id_col(c) for c in self._vc]
        self._vd=[tuple(_fmt(v,no_comma=_nc[i]) for i,v in enumerate(r))
                  for r in df[self._vc].values]
        self._render(0)

    def load_tagged(self,df,cols,tag_map):
        """tag_map: {row_index: tag_name}"""
        self.load(df,cols)
        children=self.get_children()
        for i,iid in enumerate(children):
            real_idx=self._vo+i
            if real_idx in tag_map:
                self.item(iid,tags=(tag_map[real_idx],))

    def _render(self,offset):
        for it in self.get_children(): self.delete(it)
        end=min(offset+self.PAGE,len(self._vd))
        for i,row in enumerate(self._vd[offset:end]):
            self.insert("","end",values=row,tags=("odd" if i%2 else "even",))
        self._vo=offset

    def _hclick(self,e):
        if self.identify_region(e.x,e.y)!="heading": return
        idx=int(self.identify_column(e.x)[1:])-1
        if idx>=len(self._vc): return
        col=self._vc[idx]
        self._sa=not self._sa if col==self._sc else True
        self._sc=col
        def _k(t):
            v=t[idx]
            try: return(0,float(v.replace(",","")))
            except: return(1,str(v))
        self._vd.sort(key=_k,reverse=not self._sa)
        for c in self._vc:
            self.heading(c,text=c+(" ^" if c==self._sc and self._sa
                                   else " v" if c==self._sc else ""))
        self._render(0)

    @property
    def total(self): return len(self._vd)
    @property
    def offset(self): return self._vo

# 콤마 없이 표시할 ID성 컬럼 키워드
_NO_COMMA_KW = (
    "계약번호","서비스번호","청약번호","고객번호","회원번호",
    "주문번호","신청번호","접수번호","일련번호","번호",
    "contract","service","application","order","serial","no","id",
)

def _is_id_col(col_name):
    c = str(col_name).lower().replace(" ","").replace("_","")
    return any(kw in c for kw in _NO_COMMA_KW)

def _fmt(v, no_comma=False):
    if v is None or (isinstance(v, float) and np.isnan(v)): return ""
    if no_comma:
        if isinstance(v, float) and v == int(v): return str(int(v))
        if isinstance(v, (int, np.integer)):     return str(int(v))
        if isinstance(v, float):                 return str(v)
        return str(v)
    if isinstance(v, float) and v == int(v): return f"{int(v):,}"
    if isinstance(v, (int, np.integer)):     return f"{v:,}"
    if isinstance(v, float):                 return f"{v:,.2f}"
    return str(v)

# ============================================================
# ENGINE 2: Smart Header Detector
# ============================================================
class HDet:
    @staticmethod
    def detect(xl,sh,n=10):
        try: df=xl.parse(sh,header=None,nrows=n)
        except: return 0
        scores=[]
        for i in range(min(n,len(df))):
            row=df.iloc[i]; nn=row.notna().sum()
            if nn==0: scores.append(-999); continue
            sr=sum(isinstance(v,str) for v in row if pd.notna(v))/nn
            ur=row.nunique()/nn
            nr=sum(isinstance(v,(int,float)) and not isinstance(v,bool)
                   for v in row if pd.notna(v))/nn
            scores.append(sr*2+ur*1.5-nr*2)
        return int(np.argmax(scores))

# ============================================================
# ENGINE 3: Async Loader
# ============================================================
class Loader:
    def __init__(self,done,err,prog):
        self.done=done;self.err=err;self.prog=prog

    def load(self,path,enc="auto"):
        threading.Thread(target=self._run,args=(path,enc),daemon=True).start()

    def _run(self,path,enc):
        try:
            ext=os.path.splitext(path)[1].lower()
            self.prog(5,"파일 여는 중...")
            # ── Windows secured Excel: try COM first ──────────
            if IS_WIN and ext in (".xlsx",".xls",".xlsm"):
                raw=self._win_com(path,ext,enc)
                if raw: return
            # ── normal pandas ─────────────────────────────────
            if ext in (".xlsx",".xlsm"):
                xl=pd.ExcelFile(path,engine="openpyxl")
            elif ext==".xls":
                xl=pd.ExcelFile(path,engine="xlrd")
            elif ext==".csv":
                self.prog(50,"CSV 파싱 중...")
                df=self._csv(path,enc,0)
                self.done(None,{"시트1":df},path,ext)
                return
            else:
                raise ValueError(f"지원하지 않는 형식: {ext}")
            data={}
            for i,sh in enumerate(xl.sheet_names):
                self.prog(5+int(85*i/len(xl.sheet_names)),f"시트 로딩: {sh}")
                data[sh]=xl.parse(sh,header=0)
            self.prog(95,"완료")
            self.done(xl,data,path,ext)
        except Exception as e:
            self.err(str(e))

    def _win_com(self,path,ext,enc):
        """Windows COM automation — reads password-protected / DRM files"""
        try:
            import win32com.client as w32
            xl_app=w32.Dispatch("Excel.Application")
            xl_app.Visible=False; xl_app.DisplayAlerts=False
            wb=xl_app.Workbooks.Open(os.path.abspath(path),
                                      Password="",WriteResPassword="",
                                      ReadOnly=True,IgnoreReadOnlyRecommended=True)
            tmp=path+"_UNLOCKED_.xlsx"
            wb.SaveAs(tmp,51)   # 51 = xlOpenXMLWorkbook
            wb.Close(False); xl_app.Quit()
            xl=pd.ExcelFile(tmp,engine="openpyxl")
            data={}
            for sh in xl.sheet_names: data[sh]=xl.parse(sh,header=0)
            try: os.remove(tmp)
            except: pass
            self.prog(95,"COM 로드 완료")
            self.done(xl,data,path,ext)
            return True
        except Exception:
            return False   # fall back to pandas

    @staticmethod
    def _csv(path,enc,hrow):
        encs=(["utf-8-sig","euc-kr","cp949","utf-8","latin1"]
              if enc in("auto","자동감지","자동") else [enc])
        for e in encs:
            try: return pd.read_csv(path,encoding=e,header=hrow,low_memory=False)
            except: pass
        raise ValueError("CSV 인코딩 감지 실패")

# ============================================================
# ENGINE 4: Filter Engine  (vectorised, regex, empty-check)
# ============================================================
class FEng:
    @staticmethod
    def run(df,conds):
        mask=pd.Series(True,index=df.index)
        for col,op,val in conds:
            if col not in df.columns: continue
            if op in("emp","nemp"):
                s=df[col].astype(str).str.strip()
                mask&=(s.isin(["","nan","None"]) if op=="emp"
                       else ~s.isin(["","nan","None"])); continue
            if val=="": continue
            s=df[col].astype(str)
            if   op=="ct":  m=s.str.contains(val,case=False,na=False)
            elif op=="nct": m=~s.str.contains(val,case=False,na=False)
            elif op=="eq":  m=s.str.lower()==val.lower()
            elif op=="ne":  m=s.str.lower()!=val.lower()
            elif op=="sw":  m=s.str.lower().str.startswith(val.lower(),na=False)
            elif op=="ew":  m=s.str.lower().str.endswith(val.lower(),na=False)
            elif op=="rx":
                try: m=s.str.contains(val,regex=True,na=False)
                except: m=pd.Series(True,index=df.index)
            else:
                try:
                    nv=float(val); num=pd.to_numeric(df[col],errors="coerce")
                    if   op=="gt":  m=num>nv
                    elif op=="lt":  m=num<nv
                    elif op=="gte": m=num>=nv
                    elif op=="lte": m=num<=nv
                    else: m=pd.Series(True,index=df.index)
                except: m=pd.Series(True,index=df.index)
            mask&=m
        return df[mask].copy()

# ============================================================
# ENGINE 5: Chart Engine
# ============================================================
class CEng:
    @staticmethod
    def ax(ax):
        ax.set_facecolor(C["card"])
        ax.spines[["top","right"]].set_visible(False)
        for sp in ax.spines.values(): sp.set_edgecolor(C["border"])
        ax.tick_params(colors=C["muted"],labelsize=8)
        ax.yaxis.label.set_color(C["muted"]); ax.xaxis.label.set_color(C["muted"])
        ax.title.set_color(C["text"])

    @classmethod
    def bar(cls,ax,lbl,vals,cols,horiz=False):
        cls.ax(ax); short=[str(l)[:13] for l in lbl]
        if horiz:
            bars=ax.barh(short,vals,color=cols,height=0.65)
            ax.tick_params(axis="y",colors=C["text2"],labelsize=8)
            for b,v in zip(bars,vals):
                ax.text(max(v,0),b.get_y()+b.get_height()/2,
                        f"  {v:,.0f}",va="center",fontsize=7,color=C["text2"])
        else:
            x=range(len(lbl))
            bars=ax.bar(x,vals,color=cols,width=0.68,zorder=2)
            ax.set_xticks(list(x))
            ax.set_xticklabels(short,rotation=35,ha="right",fontsize=8,color=C["text2"])
            ax.grid(axis="y",color=C["border"],alpha=0.35,zorder=0)
            for b,v in zip(bars,vals):
                ax.text(b.get_x()+b.get_width()/2,
                        b.get_height()+abs(b.get_height())*0.02,
                        f"{v:,.0f}",ha="center",fontsize=7,color=C["text2"])

    @classmethod
    def pie(cls,ax,lbl,vals,cols,donut=False):
        short=[str(l)[:12] for l in lbl]
        kw=dict(labels=short,colors=cols,autopct="%1.1f%%",pctdistance=0.78,
                wedgeprops={"linewidth":1.5,"edgecolor":C["card"]})
        if donut: kw["wedgeprops"]["width"]=0.55
        _,texts,autos=ax.pie(vals,**kw)
        for t in texts+autos: t.set_color(C["text2"]); t.set_fontsize(8)

    @classmethod
    def line(cls,ax,lbl,vals,col):
        cls.ax(ax); x=range(len(lbl))
        ax.plot(list(x),vals,color=col,linewidth=2.2,marker="o",markersize=5,zorder=3)
        ax.fill_between(list(x),vals,alpha=0.12,color=col)
        ax.set_xticks(list(x))
        ax.set_xticklabels([str(l)[:12] for l in lbl],
                           rotation=35,ha="right",fontsize=8,color=C["text2"])
        ax.grid(color=C["border"],alpha=0.3)

# ============================================================
# MAIN APP
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 분석기 PRO  v5.1")
        self.geometry("1480x900"); self.minsize(1040,640)
        self.configure(bg=C["bg"])

        # state
        self.raw_xl=None; self.raw_ext=""; self.raw_path=""
        self.wb_data={}; self.df_raw=None; self.df_view=None
        self.all_cols=[]; self.vis_cols=[]; self.col_vars={}
        self.cond_rows=[]; self._cur_fig=None
        self._pivot_df=None; self._stats_df=None
        self._qjob=None; self._recent=_load_recent()
        self._undo_stack=[]
        self._redo_stack=[]
        self._filter_presets={}
        self._preset_file=os.path.join(
            os.path.expanduser('~'),'.excel_pro_presets.json')
        self._col_order=[]
        self._load_presets()

        self._loader=Loader(
            done    =self._load_done,
            err     =lambda e:self.after(0,lambda:messagebox.showerror("오류",e)),
            prog    =lambda p,t:self.after(0,lambda:self._prog(p,t))
        )

        self._sty(); self._ui()
        self.protocol("WM_DELETE_WINDOW",self._close)

        # ── 키보드 단축키 ─────────────────────────────────────────
        mod = "Command" if IS_MAC else "Control"
        self.bind(f"<{mod}-o>", lambda e: self.open_file())
        self.bind(f"<{mod}-z>", lambda e: self._undo())
        self.bind(f"<{mod}-y>", lambda e: self._redo())
        self.bind(f"<{mod}-Shift-z>", lambda e: self._redo())
        if IS_MAC:
            self.createcommand("tk::mac::Quit", self._close)

    # ======================== STYLE ========================
    def _sty(self):
        s=ttk.Style(self)
        try: s.theme_use("clam")
        except tk.TclError: s.theme_use("default")
        s.configure(".",background=C["bg"],foreground=C["text"],
                    font=(FN,10))
        s.configure("TFrame",background=C["bg"])
        s.configure("Card.TFrame",background=C["card"])
        s.configure("TLabel",background=C["bg"],foreground=C["text"])
        s.configure("Card.TLabel",background=C["card"],foreground=C["text"])
        btns=[("TButton",C["accent"],C["bg"],C["accent_h"]),
              ("Green.TButton",C["green"],C["bg"],C["green_h"]),
              ("Purple.TButton",C["purple"],C["bg"],C["purple_h"]),
              ("Orange.TButton",C["orange"],C["bg"],C["orange_h"]),
              ("Teal.TButton",C["teal"],C["bg"],C["cyan"]),
              ("Flat.TButton",C["card2"],C["muted"],C["border"]),
              ("Sm.TButton",C["surface"],C["muted"],C["card"])]
        for nm,bg,fg,hv in btns:
            s.configure(nm,background=bg,foreground=fg,borderwidth=0,
                        focusthickness=0,font=(FN,
                                               9,"bold"),padding=(6,4))
            s.map(nm,background=[("active",hv)])
        s.configure("Treeview",background=C["card"],foreground=C["text2"],
                    fieldbackground=C["card"],rowheight=26 if IS_MAC else 23,
                    font=(FN,9))
        s.configure("Treeview.Heading",background=C["surface"],foreground=C["muted"],
                    font=(FN,9,"bold"),relief="flat")
        s.map("Treeview",background=[("selected","#1f3a5f")])
        for w in("TCombobox","TEntry","TSpinbox"):
            s.configure(w,fieldbackground=C["card2"],background=C["card2"],
                        foreground=C["text"],insertcolor=C["text"],
                        selectbackground=C["accent"],arrowcolor=C["muted"],
                        bordercolor=C["border"])
            s.map(w,fieldbackground=[("readonly",C["card2"])])
        s.configure("TNotebook",background=C["bg"],tabmargins=[2,2,0,0])
        s.configure("TNotebook.Tab",background=C["surface"],foreground=C["muted"],
                    padding=[14,7],font=(FN,10))
        s.map("TNotebook.Tab",background=[("selected",C["accent"])],
              foreground=[("selected",C["bg"])])
        s.configure("TScrollbar",background=C["border"],troughcolor=C["bg"],
                    arrowcolor=C["muted"],borderwidth=0,relief="flat")
        s.configure("Horizontal.TProgressbar",troughcolor=C["surface"],
                    background=C["accent"],borderwidth=0,thickness=4)

    # ======================== MAIN UI ========================
    def _ui(self):
        self._hdr()
        # progress
        pf=tk.Frame(self,bg=C["bg"],height=4); pf.pack(fill="x"); pf.pack_propagate(False)
        self._pbv=tk.IntVar(value=0)
        ttk.Progressbar(pf,variable=self._pbv,style="Horizontal.TProgressbar",
                        maximum=100).pack(fill="x",expand=True)
        body=tk.PanedWindow(self,orient="horizontal",bg=C["border"],sashwidth=5,sashrelief="flat")
        body.pack(fill="both",expand=True)
        sb=tk.Frame(body,bg=C["surface"],width=256); sb.pack_propagate(False)
        body.add(sb,minsize=200); self._sidebar(sb)
        main=tk.Frame(body,bg=C["bg"]); body.add(main,minsize=660)
        self._main(main)
        self._statbar()

    def _hdr(self):
        h=tk.Frame(self,bg=C["surface"],height=50); h.pack(fill="x"); h.pack_propagate(False)
        tk.Label(h,text="  Excel 분석기 PRO",bg=C["surface"],fg=C["text"],
                 font=(FN,13,"bold")).pack(side="left")
        tk.Label(h,text="v5.0",bg=C["surface"],fg=C["muted"],
                 font=(FN,8)
                 ).pack(side="left",padx=(2,16),pady=14)
        os_label=f"[{'윈도우' if IS_WIN else '맥OS' if IS_MAC else '리눅스'}]"
        tk.Label(h,text=os_label,bg=C["surface"],fg=C["cyan"],
                 font=(FN,8)).pack(side="left")
        self.lbl_file=tk.Label(h,text="  파일을 열어주세요",bg=C["surface"],
                               fg=C["muted"],font=(FN,9))
        self.lbl_file.pack(side="left",padx=8)
        _open_lbl = "  파일 열기  (Cmd+O)" if IS_MAC else "  파일 열기  (Ctrl+O)"
        ttk.Button(h,text=_open_lbl,command=self.open_file).pack(side="right",padx=6,pady=9)
        ttk.Button(h,text=" 최근 파일",style="Flat.TButton",
                   command=self._show_recent).pack(side="right",padx=2,pady=9)
        ttk.Button(h,text=" 열린 엑셀 불러오기",style="Green.TButton",
                   command=self._load_open_excel).pack(side="right",padx=6,pady=9)
        ttk.Button(h,text="↺",style="Flat.TButton",width=2,
                   command=self._undo).pack(side="right",padx=0,pady=9)
        ttk.Button(h,text="↻",style="Flat.TButton",width=2,
                   command=self._redo).pack(side="right",padx=2,pady=9)
        tk.Label(h,text="Undo/Redo",bg=C["surface"],fg=C["muted"],
                 font=(FN,8)
                 ).pack(side="right",padx=(8,0))

    def _statbar(self):
        sb=tk.Frame(self,bg=C["surface"],height=26); sb.pack(fill="x",side="bottom")
        sb.pack_propagate(False)
        self.lbl_st=tk.Label(sb,text="● 준비",bg=C["surface"],fg=C["muted"],
                             font=(FN,8))
        self.lbl_st.pack(side="left",padx=12)
        self.lbl_st2=tk.Label(sb,text="",bg=C["surface"],fg=C["muted"],
                              font=(FN,8))
        self.lbl_st2.pack(side="right",padx=12)

    def _st(self,t,fg=None):
        self.lbl_st.config(text=t,fg=fg or C["muted"]); self.update_idletasks()

    def _prog(self,v,t=""):
        self._pbv.set(v)
        if t: self._st(t,C["orange"])
        if v>=100: self.after(700,lambda:self._pbv.set(0))

    # ======================== SIDEBAR ========================
    def _sidebar(self,parent):
        cv=tk.Canvas(parent,bg=C["surface"],highlightthickness=0)
        sb2=ttk.Scrollbar(parent,orient="vertical",command=cv.yview)
        cv.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right",fill="y"); cv.pack(fill="both",expand=True)
        inn=tk.Frame(cv,bg=C["surface"])
        win=cv.create_window((0,0),window=inn,anchor="nw")
        def _cfg(e):
            cv.configure(scrollregion=cv.bbox("all"))
            cv.itemconfig(win,width=cv.winfo_width())
        inn.bind("<Configure>",_cfg)
        cv.bind("<Configure>",lambda e:cv.itemconfig(win,width=e.width))
        # 전역 바인딩 대신 사이드바 캔버스 한정 스크롤 (다른 위젯 방해 방지)
        def _mw(e):
            if IS_MAC: cv.yview_scroll(-1*e.delta,"units")
            else: cv.yview_scroll(-1*(e.delta//120),"units")
        cv.bind("<MouseWheel>", _mw)
        cv.bind("<Button-4>", lambda e: cv.yview_scroll(-1,"units"))
        cv.bind("<Button-5>", lambda e: cv.yview_scroll(1,"units"))
        inn.bind("<MouseWheel>", _mw)
        p=inn; pad={"padx":12,"pady":3}

        self._sh(p,"시트 선택"); self.cmb_sheet=ttk.Combobox(p,state="readonly")
        self.cmb_sheet.pack(fill="x",**pad)
        self.cmb_sheet.bind("<<ComboboxSelected>>",lambda e:self._sheet_chg())

        self._sh(p,"헤더 행  (자동감지 지원)")
        # 1행: 행 번호 + 적용
        hf=tk.Frame(p,bg=C["surface"]); hf.pack(fill="x",**pad)
        tk.Label(hf,text="행 번호:",bg=C["surface"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.spn=tk.Spinbox(hf,from_=0,to=50,width=5,bg=C["card2"],fg=C["text"],
                            insertbackground=C["text"],buttonbackground=C["border"],
                            relief="flat",font=(FN,9),
                            command=self._prev_hdr)
        self.spn.delete(0,"end"); self.spn.insert(0,"0"); self.spn.pack(side="left",padx=(4,6))
        ttk.Button(hf,text=" 적용",style="Sm.TButton",
                   command=self._apply_hdr).pack(side="left",fill="x",expand=True)
        # 2행: 자동감지 (전체 너비)
        hf2=tk.Frame(p,bg=C["surface"]); hf2.pack(fill="x",padx=12,pady=(2,0))
        ttk.Button(hf2,text=" AI 헤더 자동감지",style="Flat.TButton",
                   command=self._ai_hdr).pack(fill="x")
        self.lbl_hp=tk.Label(p,text="",bg=C["surface"],fg=C["cyan"],
                             font=(FN,8),
                             wraplength=210,justify="left")
        self.lbl_hp.pack(fill="x",padx=12,pady=(2,2))

        self._sh(p,"인코딩  (CSV 전용)")
        self.cmb_enc=ttk.Combobox(p,state="readonly",
            values=["자동감지","utf-8-sig","utf-8","euc-kr","cp949","latin1"])
        self.cmb_enc.set("자동감지"); self.cmb_enc.pack(fill="x",**pad)
        tk.Label(p,text="※ 한글 깨짐 시 euc-kr 선택",bg=C["surface"],fg=C["muted2"],
                 font=(FN,8)).pack(anchor="w",padx=12)

        tk.Frame(p,bg=C["border"],height=1).pack(fill="x",padx=10,pady=8)
        self._sh(p,"컬럼 선택")
        bf=tk.Frame(p,bg=C["surface"]); bf.pack(fill="x",**pad)
        ttk.Button(bf,text="전체선택",style="Sm.TButton",
                   command=lambda:self._col_all(True)).pack(side="left",expand=True,fill="x")
        ttk.Button(bf,text="전체해제",style="Sm.TButton",
                   command=lambda:self._col_all(False)).pack(side="left",expand=True,fill="x",padx=(4,0))
        bf2=tk.Frame(p,bg=C["surface"]); bf2.pack(fill="x",**pad)
        ttk.Button(bf2,text="↑ 위로",style="Sm.TButton",
                   command=lambda:self._col_move(-1)).pack(side="left",expand=True,fill="x")
        ttk.Button(bf2,text="↓ 아래",style="Sm.TButton",
                   command=lambda:self._col_move(1)).pack(side="left",expand=True,fill="x",padx=(4,0))
        self.col_frame=tk.Frame(p,bg=C["surface"]); self.col_frame.pack(fill="x",padx=12,pady=4)

        tk.Frame(p,bg=C["border"],height=1).pack(fill="x",padx=10,pady=8)
        self._sh(p,"내보내기")
        ttk.Button(p,text=" CSV 저장",style="Sm.TButton",
                   command=lambda:self._exp("csv")).pack(fill="x",padx=12,pady=2)
        ttk.Button(p,text=" XLSX 저장",style="Green.TButton",
                   command=lambda:self._exp("xlsx")).pack(fill="x",padx=12,pady=2)
        ttk.Button(p,text=" 필터 결과만",style="Flat.TButton",
                   command=self._exp_filt).pack(fill="x",padx=12,pady=2)

        tk.Frame(p,bg=C["border"],height=1).pack(fill="x",padx=10,pady=8)
        self._sh(p,"데이터 현황")
        self.lbl_info=tk.Label(p,text="파일을 열면\n데이터 정보가\n표시됩니다",
                               bg=C["surface"],fg=C["muted2"],
                               font=(FN,8),
                               justify="left")
        self.lbl_info.pack(anchor="w",padx=12,pady=2)

    def _sh(self,p,t):
        tk.Label(p,text=t,bg=C["surface"],fg=C["muted"],
                 font=(FN,8,"bold")
                 ).pack(anchor="w",padx=12,pady=(10,2))

    # ======================== MAIN TABS ========================
    def _main(self,parent):
        self.nb=ttk.Notebook(parent)
        self.nb.pack(fill="both",expand=True,padx=6,pady=6)
        self._t_filter()
        self._t_merge()
        self._t_match()
        self._t_replace()
        self._t_inspect()
        self._t_viz()
        self._t_pivot()
        self._t_stats()
        self._t_geocode()
        self._t_write()
        self._t_multmerge()

    # ─────────────────────────────────────────────────────────────
    # TAB 1: Filter / Data
    # ─────────────────────────────────────────────────────────────
    def _t_filter(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  필터 / 데이터  ")
        # filter card
        fc=tk.Frame(tab,bg=C["card"]); fc.pack(fill="x",padx=6,pady=(6,3))
        top=tk.Frame(fc,bg=C["card"]); top.pack(fill="x",padx=12,pady=(8,4))
        tk.Label(top,text="다중 조건 필터",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")).pack(side="left")
        tk.Label(top,text="전체 검색:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)
                 ).pack(side="left",padx=(20,4))
        self.var_qs=tk.StringVar(); self.var_qs.trace("w",lambda *_:self._qs())
        ttk.Entry(top,textvariable=self.var_qs,width=28).pack(side="left")
        self.cond_frame=tk.Frame(fc,bg=C["card"]); self.cond_frame.pack(fill="x",padx=12)
        bf=tk.Frame(fc,bg=C["card"]); bf.pack(fill="x",padx=12,pady=(6,10))
        ttk.Button(bf,text="＋ 조건 추가",style="Flat.TButton",command=self._add_cond).pack(side="left")
        ttk.Button(bf,text="적용",command=self._do_filter).pack(side="left",padx=6)
        ttk.Button(bf,text="↺ 초기화",style="Flat.TButton",command=self._reset_filter).pack(side="left")
        ttk.Button(bf,text=" 조건 저장",style="Flat.TButton",
                   command=self._save_preset).pack(side="left",padx=6)
        self.cmb_preset=ttk.Combobox(bf,state="readonly",width=14)
        self.cmb_preset.pack(side="left")
        ttk.Button(bf,text="불러오기",style="Flat.TButton",
                   command=self._load_preset).pack(side="left",padx=2)
        ttk.Button(bf,text="삭제",style="Flat.TButton",
                   command=self._del_preset).pack(side="left")
        self.lbl_fr=tk.Label(bf,text="",bg=C["card"],fg=C["cyan"],
                             font=(FN,9))
        self.lbl_fr.pack(side="left",padx=10)
        # table card
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        hf2=tk.Frame(tc,bg=C["card"]); hf2.pack(fill="x",padx=10,pady=(8,4))
        self.lbl_tbl=tk.Label(hf2,text="데이터 미리보기",bg=C["card"],fg=C["text"],
                              font=(FN,10,"bold"))
        self.lbl_tbl.pack(side="left")
        self.lbl_pg=tk.Label(hf2,text="",bg=C["card"],fg=C["muted"],
                             font=(FN,8))
        self.lbl_pg.pack(side="right",padx=6)
        ttk.Button(hf2,text=">>",style="Sm.TButton",command=self._pnext).pack(side="right")
        ttk.Button(hf2,text="<<",style="Sm.TButton",command=self._pprev).pack(side="right",padx=2)
        tw=tk.Frame(tc,bg=C["card"]); tw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.tree=VTree(tw,show="headings",selectmode="extended")
        vsb=ttk.Scrollbar(tw,orient="vertical",command=self.tree.yview)
        hsb=ttk.Scrollbar(tw,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        hsb.pack(side="bottom",fill="x"); vsb.pack(side="right",fill="y")
        self.tree.pack(fill="both",expand=True)
        self.tree.bind("<Control-c>",self._copy)
        self.tree.bind("<Button-3>",self._col_ctx_menu)

    # ─────────────────────────────────────────────────────────────
    # TAB 2: Merge (계약번호 기준)
    # ─────────────────────────────────────────────────────────────
    def _t_merge(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  병합  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="계약번호 기준 병합  (외부 파일)",
                 bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,6))

        # external file row
        ef=tk.Frame(cc,bg=C["card"]); ef.pack(fill="x",padx=12,pady=2)
        tk.Label(ef,text="외부 파일:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.lbl_mf=tk.Label(ef,text="(없음)",bg=C["card"],fg=C["cyan"],
                             font=(FN,9))
        self.lbl_mf.pack(side="left",padx=8)
        ttk.Button(ef,text=" 찾아보기",style="Sm.TButton",
                   command=self._load_merge_file).pack(side="left")

        # key columns
        kf=tk.Frame(cc,bg=C["card"]); kf.pack(fill="x",padx=12,pady=4)
        tk.Label(kf,text="메인 키:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.cmb_mk=ttk.Combobox(kf,state="readonly",width=18)
        self.cmb_mk.pack(side="left",padx=4)
        tk.Label(kf,text="외부 키:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left",padx=(12,0))
        self.cmb_ek=ttk.Combobox(kf,state="readonly",width=18)
        self.cmb_ek.pack(side="left",padx=4)
        tk.Label(kf,text="조인 방식:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left",padx=(12,0))
        self.cmb_join=ttk.Combobox(kf,state="readonly",width=10,
                                   values=["left","inner","outer","right"])
        self.cmb_join.set("left"); self.cmb_join.pack(side="left",padx=4)

        # columns to append
        cf2=tk.Frame(cc,bg=C["card"]); cf2.pack(fill="x",padx=12,pady=(0,4))
        tk.Label(cf2,text="추가할 컬럼 선택 (다중):",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(anchor="w")
        lbf=tk.Frame(cf2,bg=C["card"]); lbf.pack(fill="x")
        sb3=ttk.Scrollbar(lbf,orient="vertical")
        self.lb_mcols=tk.Listbox(lbf,selectmode="multiple",bg=C["card2"],fg=C["text2"],
                                 selectbackground=C["accent"],height=4,
                                 font=(FN,9),
                                 yscrollcommand=sb3.set)
        sb3.config(command=self.lb_mcols.yview)
        sb3.pack(side="right",fill="y"); self.lb_mcols.pack(fill="x",expand=True)

        bf2=tk.Frame(cc,bg=C["card"]); bf2.pack(fill="x",padx=12,pady=(0,10))
        ttk.Button(bf2,text=" 병합 실행",style="Teal.TButton",
                   command=self._do_merge).pack(side="left")
        ttk.Button(bf2,text="결과 저장",style="Green.TButton",
                   command=self._save_merged).pack(side="left",padx=6)
        self.lbl_mr=tk.Label(bf2,text="",bg=C["card"],fg=C["cyan"],
                             font=(FN,9))
        self.lbl_mr.pack(side="left",padx=8)

        # merge preview
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(tc,text="병합 미리보기",bg=C["card"],fg=C["text"],
                 font=(FN,10,"bold")
                 ).pack(anchor="w",padx=10,pady=(8,4))
        mw=tk.Frame(tc,bg=C["card"]); mw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.mtree=VTree(mw,show="headings")
        mvsb=ttk.Scrollbar(mw,orient="vertical",command=self.mtree.yview)
        mhsb=ttk.Scrollbar(mw,orient="horizontal",command=self.mtree.xview)
        self.mtree.configure(yscrollcommand=mvsb.set,xscrollcommand=mhsb.set)
        mhsb.pack(side="bottom",fill="x"); mvsb.pack(side="right",fill="y")
        self.mtree.pack(fill="both",expand=True)

        self._merge_df=None; self._ext_df=None

    # ─────────────────────────────────────────────────────────────
    # TAB 3: Replace (본부/지사명 치환)
    # ─────────────────────────────────────────────────────────────
    def _t_replace(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  치환  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="본부 / 지사명 치환",
                 bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,6))

        rf=tk.Frame(cc,bg=C["card"]); rf.pack(fill="x",padx=12,pady=2)
        tk.Label(rf,text="대상 컬럼:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.cmb_rcol=ttk.Combobox(rf,state="readonly",width=20)
        self.cmb_rcol.pack(side="left",padx=4)
        tk.Label(rf,text="치환 방식:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left",padx=(12,0))
        self.cmb_rmode=ttk.Combobox(rf,state="readonly",width=16,
            values=["완전일치","포함","정규식"])
        self.cmb_rmode.set("완전일치"); self.cmb_rmode.pack(side="left",padx=4)

        # rules table
        tk.Label(cc,text="치환 규칙  (원본 → 변경)",bg=C["card"],fg=C["muted"],
                 font=(FN,8,"bold")
                 ).pack(anchor="w",padx=12,pady=(6,2))
        rte=tk.Frame(cc,bg=C["card"]); rte.pack(fill="x",padx=12)
        rhf=tk.Frame(rte,bg=C["card"]); rhf.pack(fill="x")
        for t,w in [("원본",22),("변경",22)]:
            tk.Label(rhf,text=t,bg=C["card"],fg=C["muted"],
                     font=(FN,8),
                     width=w,anchor="w").pack(side="left",padx=2)
        self.rep_frame=tk.Frame(rte,bg=C["card"]); self.rep_frame.pack(fill="x")
        self.rep_rows=[]   # list of (from_var, to_var, frame)
        for _ in range(3): self._add_rep_row()

        bf3=tk.Frame(cc,bg=C["card"]); bf3.pack(fill="x",padx=12,pady=(6,10))
        ttk.Button(bf3,text="+ 행 추가",style="Flat.TButton",
                   command=self._add_rep_row).pack(side="left")
        ttk.Button(bf3,text=" CSV 규칙 불러오기",style="Flat.TButton",
                   command=self._load_rep_csv).pack(side="left",padx=4)
        ttk.Button(bf3,text="️ 치환 적용",style="Teal.TButton",
                   command=self._do_replace).pack(side="left",padx=4)
        ttk.Button(bf3,text="초기화",style="Flat.TButton",
                   command=self._reset_replace).pack(side="left",padx=4)
        self.lbl_rr=tk.Label(bf3,text="",bg=C["card"],fg=C["cyan"],
                             font=(FN,9))
        self.lbl_rr.pack(side="left",padx=8)

        # preview
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(tc,text="치환 미리보기",bg=C["card"],fg=C["text"],
                 font=(FN,10,"bold")
                 ).pack(anchor="w",padx=10,pady=(8,4))
        rw=tk.Frame(tc,bg=C["card"]); rw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.rtree=VTree(rw,show="headings")
        rvsb=ttk.Scrollbar(rw,orient="vertical",command=self.rtree.yview)
        rhsb=ttk.Scrollbar(rw,orient="horizontal",command=self.rtree.xview)
        self.rtree.configure(yscrollcommand=rvsb.set,xscrollcommand=rhsb.set)
        rhsb.pack(side="bottom",fill="x"); rvsb.pack(side="right",fill="y")
        self.rtree.pack(fill="both",expand=True)

    # ─────────────────────────────────────────────────────────────
    # TAB 4: Inspect (중복/공백 점검)
    # ─────────────────────────────────────────────────────────────
    def _t_inspect(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  점검  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="중복 / 공백 점검",
                 bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,6))
        cf=tk.Frame(cc,bg=C["card"]); cf.pack(fill="x",padx=12,pady=2)
        tk.Label(cf,text="중복 기준 컬럼:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        lbf2=tk.Frame(cf,bg=C["card"]); lbf2.pack(side="left",padx=8)
        isb=ttk.Scrollbar(lbf2,orient="vertical")
        self.lb_icols=tk.Listbox(lbf2,selectmode="multiple",bg=C["card2"],fg=C["text2"],
                                 selectbackground=C["accent"],height=3,
                                 font=(FN,9),
                                 yscrollcommand=isb.set)
        isb.config(command=self.lb_icols.yview)
        isb.pack(side="right",fill="y"); self.lb_icols.pack(side="left")
        bf4=tk.Frame(cc,bg=C["card"]); bf4.pack(fill="x",padx=12,pady=(0,10))
        ttk.Button(bf4,text=" 점검 실행",style="Orange.TButton",
                   command=self._do_inspect).pack(side="left")
        ttk.Button(bf4,text=" 이슈 내보내기",style="Green.TButton",
                   command=self._exp_issues).pack(side="left",padx=6)
        self.lbl_ir=tk.Label(bf4,text="",bg=C["card"],fg=C["yellow"],
                             font=(FN,9))
        self.lbl_ir.pack(side="left",padx=8)

        # KPI row
        self.kpi_frame=tk.Frame(tab,bg=C["bg"]); self.kpi_frame.pack(fill="x",padx=6,pady=4)

        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(tc,text="이슈 목록",bg=C["card"],fg=C["text"],
                 font=(FN,10,"bold")
                 ).pack(anchor="w",padx=10,pady=(8,4))
        iw=tk.Frame(tc,bg=C["card"]); iw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.itree=VTree(iw,show="headings")
        ivsb=ttk.Scrollbar(iw,orient="vertical",command=self.itree.yview)
        ihsb=ttk.Scrollbar(iw,orient="horizontal",command=self.itree.xview)
        self.itree.configure(yscrollcommand=ivsb.set,xscrollcommand=ihsb.set)
        ihsb.pack(side="bottom",fill="x"); ivsb.pack(side="right",fill="y")
        self.itree.pack(fill="both",expand=True)
        self._issues_df=None

    # ─────────────────────────────────────────────────────────────
    # TAB 5: Chart / KPI
    # ─────────────────────────────────────────────────────────────
    def _t_viz(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  차트 / KPI  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="차트 & KPI",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,4))
        ttk.Button(cc,text="  관리지사 × 계약건수 × 월정료  자동 요약",
                   style="Purple.TButton",command=self._auto_kpi
                   ).pack(fill="x",padx=12,pady=(0,6))
        mf=tk.Frame(cc,bg=C["card"]); mf.pack(fill="x",padx=12,pady=(0,10))
        fn=FN
        for lbl,attr,w in [("그룹(X)","cmb_vg",15),("집계값(Y)","cmb_vv",15),
                            ("집계방식","cmb_va",8),("차트종류","cmb_vt",12)]:
            tk.Label(mf,text=lbl,bg=C["card"],fg=C["muted"],
                     font=(fn,8)).pack(side="left",padx=(6,2))
            cmb=ttk.Combobox(mf,state="readonly",width=w)
            setattr(self,attr,cmb); cmb.pack(side="left")
        self.cmb_va["values"]=list(AGG.keys()); self.cmb_va.set("합계")
        self.cmb_vt["values"]=["막대","가로막대","파이","도넛","꺾은선","누적막대"]
        self.cmb_vt.set("막대")
        ttk.Button(mf,text=" 차트 생성",command=self._build_chart).pack(side="left",padx=(8,4))
        ttk.Button(mf,text=" 이미지 저장",style="Flat.TButton",command=self._save_chart).pack(side="left")
        ttk.Button(mf,text="↺ 초기화",style="Flat.TButton",command=self._clr_chart).pack(side="left",padx=4)
        self.kpi_vframe=tk.Frame(tab,bg=C["bg"]); self.kpi_vframe.pack(fill="x",padx=6,pady=2)
        self.chart_frame=tk.Frame(tab,bg=C["card"])
        self.chart_frame.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(self.chart_frame,text="  차트를 생성하면 여기에 표시됩니다\n\n상단의 [그룹(X)] [집계값(Y)] 설정 후 [차트 생성] 클릭",
                 bg=C["card"],fg=C["muted2"],
                 font=(FN,12)).pack(expand=True)

    # ─────────────────────────────────────────────────────────────
    # TAB 6: Pivot
    # ─────────────────────────────────────────────────────────────
    def _t_pivot(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  피벗  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="피벗 집계",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,4))
        rf=tk.Frame(cc,bg=C["card"]); rf.pack(fill="x",padx=12,pady=(0,10))
        fn=FN
        for lbl,attr,w in [("행(그룹)","cmb_pr",16),("열(선택)","cmb_pc",14),
                            ("집계값(Y)","cmb_pv",16),("집계방식","cmb_pa",8)]:
            tk.Label(rf,text=lbl,bg=C["card"],fg=C["muted"],
                     font=(fn,8)).pack(side="left",padx=(6,2))
            cmb=ttk.Combobox(rf,state="readonly",width=w)
            setattr(self,attr,cmb); cmb.pack(side="left")
        self.cmb_pa["values"]=list(AGG.keys()); self.cmb_pa.set("합계")
        ttk.Button(rf,text=" 집계 실행",command=self._pivot).pack(side="left",padx=(10,4))
        ttk.Button(rf,text=" 저장",style="Green.TButton",command=self._exp_pivot).pack(side="left")
        ttk.Button(rf,text=" 차트",style="Flat.TButton",command=self._pivot_chart).pack(side="left",padx=4)
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        self.lbl_pv=tk.Label(tc,text="집계 결과",bg=C["card"],fg=C["text"],
                             font=(FN,10,"bold"))
        self.lbl_pv.pack(anchor="w",padx=10,pady=(8,4))
        pw=tk.Frame(tc,bg=C["card"]); pw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.ptree=ttk.Treeview(pw,show="headings")
        pvsb=ttk.Scrollbar(pw,orient="vertical",command=self.ptree.yview)
        phsb=ttk.Scrollbar(pw,orient="horizontal",command=self.ptree.xview)
        self.ptree.configure(yscrollcommand=pvsb.set,xscrollcommand=phsb.set)
        self.ptree.tag_configure("odd",background="#131820")
        self.ptree.tag_configure("total",background="#1f3650",foreground=C["yellow"])
        phsb.pack(side="bottom",fill="x"); pvsb.pack(side="right",fill="y")
        self.ptree.pack(fill="both",expand=True)

    # ─────────────────────────────────────────────────────────────
    # TAB 7: Statistics
    # ─────────────────────────────────────────────────────────────
    def _t_stats(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  통계  ")
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="기초 통계 분석",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,4))
        bf=tk.Frame(cc,bg=C["card"]); bf.pack(fill="x",padx=12,pady=(0,10))
        ttk.Button(bf,text=" 통계 계산",command=self._stats).pack(side="left")
        ttk.Button(bf,text="Save",style="Green.TButton",
                   command=self._exp_stats).pack(side="left",padx=6)
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        sw=tk.Frame(tc,bg=C["card"]); sw.pack(fill="both",expand=True,padx=8,pady=(4,8))
        self.stree=ttk.Treeview(sw,show="headings")
        svsb=ttk.Scrollbar(sw,orient="vertical",command=self.stree.yview)
        shsb=ttk.Scrollbar(sw,orient="horizontal",command=self.stree.xview)
        self.stree.configure(yscrollcommand=svsb.set,xscrollcommand=shsb.set)
        self.stree.tag_configure("odd",background="#131820")
        shsb.pack(side="bottom",fill="x"); svsb.pack(side="right",fill="y")
        self.stree.pack(fill="both",expand=True)

    # ======================== OPEN EXCEL LOAD ========================
    def _get_open_books(self):
        """열린 Excel 파일 목록 반환 (Mac: AppleScript 우선 → xlwings, Win: xlwings → win32com)"""
        books = []

        # ── Mac: AppleScript (가장 신뢰성 높음) ──────────────────
        if IS_MAC:
            try:
                script = (
                    'tell application "Microsoft Excel"\n'
                    '  set out to ""\n'
                    '  repeat with wb in workbooks\n'
                    '    set wn to name of wb\n'
                    '    try\n'
                    '      set wp to full name of wb as string\n'
                    '    on error\n'
                    '      set wp to ""\n'
                    '    end try\n'
                    '    set out to out & wn & "|||" & wp & "\n"\n'
                    '  end repeat\n'
                    '  return out\n'
                    'end tell'
                )
                r = subprocess.run(['osascript','-e',script],
                                   capture_output=True,text=True,timeout=6)
                if r.returncode==0 and r.stdout.strip():
                    # xlwings 객체 맵 (시트 목록용)
                    xw_map = {}
                    try:
                        import xlwings as _xwt
                        for _a in _xwt.apps:
                            for _w in _a.books:
                                xw_map[_w.name] = _w
                    except Exception: pass

                    for line in r.stdout.strip().split('\n'):
                        if '|||' not in line: continue
                        name, path = line.split('|||',1)
                        name=name.strip(); path=path.strip()
                        # Mac 콜론 경로 → POSIX 변환
                        if path and ':' in path and not path.startswith('/'):
                            try:
                                pr=subprocess.run(
                                    ['osascript','-e',f'POSIX path of "{path}"'],
                                    capture_output=True,text=True,timeout=3)
                                if pr.returncode==0: path=pr.stdout.strip()
                            except Exception: pass
                        wb_obj = xw_map.get(name)
                        sheets = []
                        if wb_obj:
                            try: sheets=[s.name for s in wb_obj.sheets]
                            except Exception: pass
                        if not sheets and path and os.path.exists(path):
                            try:
                                sheets=pd.ExcelFile(path,engine='openpyxl').sheet_names
                            except Exception: sheets=['Sheet1']
                        if not sheets: sheets=['Sheet1']
                        books.append({"source":"applescript","app":None,
                                      "wb":wb_obj,"name":name,"path":path,
                                      "sheets":sheets})
            except Exception: pass

        # ── xlwings (Mac fallback / Win primary) ─────────────────
        if not books:
            try:
                import xlwings as xw
                for _a in xw.apps:
                    for _w in _a.books:
                        try:
                            books.append({"source":"xlwings","app":_a,"wb":_w,
                                          "name":_w.name,"path":_w.fullname,
                                          "sheets":[s.name for s in _w.sheets]})
                        except Exception: pass
            except Exception: pass

        # ── Win32COM (Windows 최후 수단) ──────────────────────────
        if IS_WIN and not books:
            try:
                import win32com.client as _c
                _xl=_c.GetActiveObject("Excel.Application")
                for _w in _xl.Workbooks:
                    try:
                        books.append({"source":"win32com","app":None,"wb":None,
                                      "_w32wb":_w,"name":_w.Name,"path":_w.FullName,
                                      "sheets":[_w.Sheets(i+1).Name
                                                for i in range(_w.Sheets.Count)]})
                    except Exception: pass
            except Exception: pass

        return books

    def _load_open_excel(self):
        """열려 있는 엑셀 파일 목록 감지 → 선택 팝업 → 읽기"""
        import os, threading as _th, traceback as _tb
        import tempfile as _tf, shutil as _sh
        fn = FN

        # ── STEP 1. 열린 파일 목록 수집 ───────────────────────────
        books = self._get_open_books()

        if not books:
            messagebox.showwarning("열린 엑셀 없음",
                "Excel 파일이 열려 있지 않습니다.\n\n"
                "1. Excel 에서 파일을 열어주세요\n"
                "2. 버튼을 다시 클릭하세요\n"
                "3. 그래도 안 되면 [파일 열기]를 사용하세요")
            return

        # ── STEP 2. 팝업 UI 구성 ──────────────────────────────────
        pop = tk.Toplevel(self)
        pop.title("열린 Excel 파일 선택")
        pop.configure(bg=C["surface"])
        pop.geometry("700x560")
        pop.resizable(False, True)
        pop.transient(self); pop.grab_set(); pop.lift(); pop.focus_force()

        tk.Label(pop, text="현재 열려 있는 Excel 파일",
                 bg=C["surface"], fg=C["text"],
                 font=(fn,12,"bold")).pack(pady=(14,2))
        tk.Label(pop, text=f"총 {len(books)}개 파일  —  파일·시트 선택 후 [불러오기]",
                 bg=C["surface"], fg=C["muted"], font=(fn,9)).pack(pady=(0,8))

        body = tk.Frame(pop, bg=C["surface"])
        body.pack(fill="both", expand=True, padx=16)

        sel_key   = tk.StringVar()
        sel_sheet = tk.StringVar()
        sel_hdr   = tk.IntVar(value=1)
        wb_map    = {}

        # 파일 목록
        tk.Label(body, text="파일 목록", bg=C["surface"], fg=C["muted"],
                 font=(fn,8,"bold")).pack(anchor="w", pady=(0,3))
        list_h = min(130, len(books)*36+8)
        lf = tk.Frame(body, bg=C["card"], height=list_h)
        lf.pack(fill="x"); lf.pack_propagate(False)
        lc = tk.Canvas(lf, bg=C["card"], highlightthickness=0)
        ls_sb = ttk.Scrollbar(lf, orient="vertical", command=lc.yview)
        li = tk.Frame(lc, bg=C["card"])
        lc.create_window((0,0), window=li, anchor="nw")
        lc.configure(yscrollcommand=ls_sb.set)
        li.bind("<Configure>", lambda e: lc.configure(scrollregion=lc.bbox("all")))
        ls_sb.pack(side="right", fill="y"); lc.pack(fill="both", expand=True)

        lbl_sel = tk.Label(body, text="", bg=C["surface"],
                           fg=C["cyan"], font=(fn,8), wraplength=640)
        lbl_sel.pack(anchor="w", pady=(4,0))

        # 시트 + 헤더 행
        tk.Label(body, text="시트  /  헤더 행 번호", bg=C["surface"],
                 fg=C["muted"], font=(fn,8,"bold")).pack(anchor="w", pady=(10,3))
        shr = tk.Frame(body, bg=C["surface"]); shr.pack(fill="x")
        cmb_sh = ttk.Combobox(shr, textvariable=sel_sheet,
                               state="readonly", width=22)
        cmb_sh.pack(side="left")
        tk.Label(shr, text="  헤더:", bg=C["surface"],
                 fg=C["muted"], font=(fn,9)).pack(side="left")
        tk.Spinbox(shr, from_=0, to=30, width=3, textvariable=sel_hdr,
                   bg=C["card2"], fg=C["text"], insertbackground=C["text"],
                   buttonbackground=C["border"], relief="flat", font=(fn,9)
                   ).pack(side="left", padx=2)
        btn_prev = ttk.Button(shr, text=" 미리보기", style="Sm.TButton",
                              command=lambda: _hdr_preview())
        btn_prev.pack(side="left", padx=4)
        btn_auto = ttk.Button(shr, text=" 자동감지", style="Sm.TButton",
                              command=lambda: _hdr_auto())
        btn_auto.pack(side="left", padx=2)

        # 빠른 헤더 선택 버튼들
        quick_f = tk.Frame(body, bg=C["surface"]); quick_f.pack(fill="x", pady=(2,0))
        tk.Label(quick_f, text="빠른 선택:", bg=C["surface"], fg=C["muted2"],
                 font=(fn,8)).pack(side="left")
        for _h in range(5):
            ttk.Button(quick_f, text=f"{_h}행", style="Sm.TButton",
                       command=lambda h=_h: (sel_hdr.set(h),
                                              pop.after(50, _hdr_preview))
                       ).pack(side="left", padx=2)

        lbl_hdr_prev = tk.Label(body,
            text="  헤더 행 번호 기본 1 — 자동감지는 추천만 (직접 조정 가능)",
            bg=C["surface"], fg=C["muted2"],
            font=(fn,8), wraplength=640, anchor="w")
        lbl_hdr_prev.pack(fill="x", pady=(2,4))

        # 옵션
        var_copy = tk.BooleanVar(value=True)
        tk.Checkbutton(body,
            text="임시 복사 후 읽기  (보안/편집중 파일 권장)",
            variable=var_copy,
            bg=C["surface"], fg=C["muted"], selectcolor=C["accent"],
            activebackground=C["surface"], font=(fn,8)
            ).pack(anchor="w", pady=(4,0))

        lbl_prog = tk.Label(body, text="", bg=C["surface"],
                            fg=C["orange"], font=(fn,9,"bold"))
        lbl_prog.pack(anchor="w", pady=(4,0))

        # 버튼
        btn_f = tk.Frame(pop, bg=C["surface"]); btn_f.pack(pady=10)
        btn_load = ttk.Button(btn_f, text="  불러오기",
                              command=lambda: _do_load())
        btn_load.pack(side="left", padx=6, ipadx=10, ipady=4)
        ttk.Button(btn_f, text="취소", style="Flat.TButton",
                   command=pop.destroy).pack(side="left", padx=6)
        ttk.Button(btn_f, text=" 새로고침", style="Flat.TButton",
                   command=lambda: (pop.destroy(),
                                    self.after(80, self._load_open_excel))
                   ).pack(side="left", padx=6)

        # ── STEP 3. 헬퍼 함수 정의 (버튼 이후에 정의, lambda로 지연 호출) ─

        def _hdr_preview():
            key2 = sel_key.get(); sh2 = sel_sheet.get()
            hdr2 = sel_hdr.get(); bk2 = wb_map.get(key2)
            # 선택 안 된 경우 조용히 무시 (자동 호출 타이밍 문제 방지)
            if not bk2 or not sh2:
                return
            fp2  = bk2.get("path", "")
            cols = None
            err_detail = ""

            # 방법 A: pandas 직접 읽기 (파일 경로가 유효할 때)
            try:
                import pandas as _ppd
                if os.path.exists(fp2):
                    peek = _ppd.read_excel(fp2, sheet_name=sh2,
                                           header=hdr2, nrows=0,
                                           engine="openpyxl")
                    cols = [str(c) for c in peek.columns]
            except Exception as e:
                err_detail += f"pandas: {e}  "

            # 방법 B: xlwings COM (파일 열려있을 때)
            if cols is None and bk2.get("wb"):
                try:
                    ws2 = bk2["wb"].sheets[sh2]
                    row_vals = ws2.range(
                        (hdr2+1, 1), (hdr2+1, 60)).value or []
                    cols = [str(v).strip() for v in row_vals
                            if v is not None and str(v).strip() not in
                            ("", "nan", "None")]
                except Exception as e:
                    err_detail += f"xlwings: {e}  "

            # 방법 C: win32com (A,B 모두 실패 시)
            if cols is None and IS_WIN:
                try:
                    import win32com.client as _wc2
                    _xl2 = _wc2.GetActiveObject("Excel.Application")
                    for _wb2 in _xl2.Workbooks:
                        if _wb2.Name == bk2["name"]:
                            _ws2 = _wb2.Sheets(sh2)
                            cols = []
                            for ci in range(1, 61):
                                v = _ws2.Cells(hdr2+1, ci).Value
                                if v is None: break
                                cols.append(str(v).strip())
                            break
                except Exception as e:
                    err_detail += f"win32com: {e}"

            if cols:
                cols_clean = [c for c in cols if c not in ("","nan","None")]
                preview = "  |  ".join(c[:15] for c in cols_clean[:8])
                if len(cols_clean) > 8:
                    preview += f"  ...+{len(cols_clean)-8}"
                lbl_hdr_prev.config(
                    text=f" {hdr2}행 헤더 ({len(cols_clean)}컬럼) → {preview}",
                    fg=C["cyan"])
            else:
                # 모두 실패 — 오류 없이 안내만 표시
                lbl_hdr_prev.config(
                    text=f"ℹ️ {hdr2}행 선택됨 — 불러오기 후 컬럼 확인하세요",
                    fg=C["muted"])

        def _hdr_auto():
            key2 = sel_key.get(); sh2 = sel_sheet.get()
            bk2  = wb_map.get(key2)
            if not bk2 or not sh2: return
            fp2  = bk2.get("path","")
            try:
                import pandas as _ppd, numpy as _np
                raw10 = None
                if os.path.exists(fp2):
                    try:
                        raw10 = _ppd.read_excel(fp2, sheet_name=sh2,
                                                header=None, nrows=10,
                                                engine="openpyxl")
                    except Exception: pass
                if raw10 is None and bk2.get("wb"):
                    try:
                        _ws_p = bk2["wb"].sheets[sh2]
                        _rv   = _ws_p.range((1,1),(10,50)).value or []
                        if _rv and isinstance(_rv[0], list):
                            raw10 = _ppd.DataFrame(_rv)
                        elif _rv:
                            raw10 = _ppd.DataFrame([[v] for v in _rv])
                    except Exception: pass

                if raw10 is None or len(raw10) == 0:
                    lbl_hdr_prev.config(
                        text=f"데이터 읽기 실패 — 현재 헤더 행 {sel_hdr.get()} 유지",
                        fg=C["orange"]); return

                # ── 헤더 행 점수 계산 (강화 버전) ─────────────
                n_cols = len(raw10.columns)
                scores = []
                for i in range(min(10, len(raw10))):
                    row = raw10.iloc[i]
                    nn  = row.notna().sum()
                    if nn == 0: scores.append(-999); continue

                    vals = [v for v in row if _ppd.notna(v)]
                    # 문자열 비율 (높을수록 헤더 가능성 ↑)
                    sr = sum(isinstance(v, str) for v in vals) / nn
                    # 유니크 비율 (헤더는 중복이 거의 없음)
                    ur = row.nunique() / nn
                    # 숫자 비율 (높으면 데이터행)
                    nr = sum(isinstance(v,(int,float)) and not isinstance(v,bool)
                             for v in vals) / nn
                    # 채움 비율 (헤더는 보통 컬럼 대부분을 채움)
                    fill_r = nn / max(n_cols, 1)
                    # 모든 문자열이고 길이가 짧은 값들 (헤더 특성)
                    short_str = sum(isinstance(v,str) and len(str(v)) < 30
                                    for v in vals) / max(nn, 1)
                    score = (sr*3 + ur*2 + short_str*1 +
                             fill_r*0.5 - nr*3)
                    scores.append(score)

                best = int(_np.argmax(scores))

                # 상위 2개 점수가 비슷하면 더 앞 행 선택 (보수적)
                sorted_scores = sorted(enumerate(scores),
                                       key=lambda x: x[1], reverse=True)
                if len(sorted_scores) >= 2:
                    top1_idx, top1_sc = sorted_scores[0]
                    top2_idx, top2_sc = sorted_scores[1]
                    if top2_sc > top1_sc * 0.85 and top2_idx < top1_idx:
                        best = top2_idx

                # 자동감지 결과를 버튼으로만 안내 — sel_hdr 직접 변경 안 함
                cur = sel_hdr.get()
                if best != cur:
                    lbl_hdr_prev.config(
                        text=f" 추정: {best}행  현재: {cur}행  "
                             f"— [{best}행] 버튼 클릭하면 적용됩니다",
                        fg=C["yellow"])
                else:
                    lbl_hdr_prev.config(
                        text=f" 자동감지: {best}행이 헤더로 추정  "
                             f"(현재 설정과 일치)",
                        fg=C["cyan"])
                pop.after(150, _hdr_preview)
            except Exception as e:
                lbl_hdr_prev.config(
                    text=f"자동감지 오류: {e}",
                    fg=C["orange"])

        def _on_pick(bk):
            sel_key.set(bk["_key"])
            cmb_sh["values"] = bk["sheets"]
            cmb_sh.set(bk["sheets"][0] if bk["sheets"] else "")
            lbl_sel.config(text=f"[{bk['source']}]  {bk['name']}  "
                               f"|  시트 {len(bk['sheets'])}개  |  {bk['path']}")
            # 헤더 기본값 1 유지, 자동감지 후 미리보기 자동 실행
            pop.after(300, _hdr_auto)

        cmb_sh.bind("<<ComboboxSelected>>", lambda e: pop.after(100, _hdr_auto))

        def _do_load():
            key  = sel_key.get(); sh = sel_sheet.get()
            hdr  = sel_hdr.get(); copy = var_copy.get()
            if not key:
                messagebox.showwarning("주의","파일을 선택하세요.",parent=pop); return
            if not sh:
                messagebox.showwarning("주의","시트를 선택하세요.",parent=pop); return
            bk = wb_map.get(key)
            if not bk:
                messagebox.showerror("오류","파일 정보 없음",parent=pop); return
            btn_load.config(state="disabled")
            lbl_prog.config(text="읽는 중...")
            self._st("열린 엑셀 읽는 중...", C["orange"])
            self._prog(5, "읽기 시작...")
            _th.Thread(target=_do_read, args=(bk,sh,hdr,copy), daemon=True).start()

        def _do_read(bk, sh, hdr, copy):
            try:
                import pythoncom; pythoncom.CoInitialize(); _coinit=True
            except Exception: _coinit=False

            import pandas as _pd2
            log=[]; df=None; fp=bk.get("path",""); tmp_path=None

            def _log(msg):
                log.append(msg)
                self.after(0, lambda m=msg: lbl_prog.config(text=m))

            try:
                # A. 임시복사 → pandas
                if copy:
                    try:
                        tmp_fd, tmp_path = _tf.mkstemp(suffix=".xlsx")
                        os.close(tmp_fd); ok=False
                        if bk.get("source")=="win32com" and bk.get("_w32wb"):
                            try:
                                import win32com.client as _wc
                                _xl2 = _wc.GetActiveObject("Excel.Application")
                                for _b in _xl2.Workbooks:
                                    if _b.Name==bk["name"]:
                                        _b.SaveCopyAs(tmp_path); ok=True
                                        _log("A1: win32com OK"); break
                            except Exception as e: _log(f"A1: fail → {e}")
                        if not ok and bk.get("wb"):
                            try:
                                import xlwings as _xw2
                                for _a2 in _xw2.apps:
                                    for _b2 in _a2.books:
                                        if _b2.name==bk["name"]:
                                            _b2.api.SaveCopyAs(tmp_path)
                                            ok=True; _log("A2: xlwings OK"); break
                                    if ok: break
                            except Exception as e: _log(f"A2: fail → {e}")
                        if not ok and os.path.exists(fp):
                            try:
                                _sh.copy2(fp, tmp_path); ok=True; _log("A3: shutil OK")
                            except Exception as e: _log(f"A3: fail → {e}")
                        if ok:
                            try:
                                self.after(0,lambda:self._prog(45,"pandas 읽기..."))
                                df=_pd2.read_excel(tmp_path,sheet_name=sh,
                                                   header=hdr,engine="openpyxl")
                                _log(f"A: pandas OK ({len(df)}행)")
                            except Exception as e: _log(f"A: pandas fail → {e}"); df=None
                        else: _log("A: 복사 실패 — 다음 경로 시도")
                    except Exception as e: _log(f"A: 전체 fail → {e}")
                    finally:
                        if tmp_path and os.path.exists(tmp_path):
                            try: os.remove(tmp_path)
                            except: pass

                # B. 직접 경로
                if df is None:
                    if os.path.exists(fp):
                        try:
                            self.after(0,lambda:self._prog(55,"직접경로 읽기..."))
                            ext=os.path.splitext(fp)[1].lower()
                            eng="openpyxl" if ext in(".xlsx",".xlsm") else "xlrd"
                            df=_pd2.read_excel(fp,sheet_name=sh,header=hdr,engine=eng)
                            _log(f"B: OK ({len(df)}행)")
                        except Exception as e: _log(f"B: fail → {e}"); df=None
                    else: _log(f"B: 경로 없음 → {fp}")

                # C. xlwings COM
                if df is None:
                    try:
                        import xlwings as _xw3
                        self.after(0,lambda:self._prog(65,"COM 읽기..."))
                        _ws3=None; _app3=None
                        for _a3 in _xw3.apps:
                            for _b3 in _a3.books:
                                if _b3.name==bk["name"]:
                                    _ws3=_b3.sheets[sh]; _app3=_a3; break
                            if _ws3: break
                        if _ws3 is None: _log("C: 파일 재획득 실패")
                        else:
                            try: _app3.screen_updating=False
                            except: pass
                            try: raw=_ws3.used_range.value
                            finally:
                                try: _app3.screen_updating=True
                                except: pass
                            if raw is None: _log("C: 빈 시트")
                            else:
                                if not isinstance(raw,list): raw=[[raw]]
                                elif raw and not isinstance(raw[0],list):
                                    raw=[[v] for v in raw]
                                seen4={}; cols4=[]
                                hdr_row=raw[hdr] if hdr<len(raw) else []
                                for ci,cv in enumerate(hdr_row):
                                    nm=str(cv).strip() if cv is not None else ""
                                    if nm in("","nan","None"): nm=f"_C{ci}"
                                    k4=nm; n4=seen4.get(k4,0)
                                    if n4: nm=f"{k4}_{n4}"
                                    seen4[k4]=n4+1; cols4.append(nm)
                                df=_pd2.DataFrame(raw[hdr+1:] if hdr<len(raw) else raw,
                                                  columns=cols4)
                                _log(f"C: COM OK ({len(df)}행)")
                    except Exception as e: _log(f"C: fail → {e}"); df=None

                if df is None:
                    raise RuntimeError("모든 읽기 경로 실패\n\n"+
                                       "\n".join(log))

                df=df.dropna(how="all").reset_index(drop=True)
                df.columns=[str(c).strip() if str(c) not in("nan","None","")
                            else f"_C{i}" for i,c in enumerate(df.columns)]
                self.after(0,lambda:self._prog(90,"화면 반영..."))
                self.after(0,lambda d=df,n=bk["name"],s=sh,f2=fp,h=hdr:
                    _finish_ok(d,n,s,f2,h))

            except Exception as ex:
                tb=_tb.format_exc()
                log_str="\n".join(log) if log else "없음"
                err=(f"오류: {ex}\n\n=== 경로별 결과 ===\n{log_str}"
                     f"\n\n=== Traceback ===\n{tb}")
                self.after(0,lambda t=err: _finish_err(t))
            finally:
                if _coinit:
                    try:
                        import pythoncom as _pc2; _pc2.CoUninitialize()
                    except Exception: pass

        def _finish_ok(df, name, sh, fp, hdr):
            try: pop.destroy()
            except: pass
            self._on_xw_loaded(df, name, sh, fp, hdr)

        def _finish_err(err_txt):
            btn_load.config(state="normal")
            lbl_prog.config(text="읽기 실패 — 오류창 확인", fg=C["red"])
            self._st("읽기 실패", C["red"]); self._prog(0)
            ew = tk.Toplevel(pop)
            ew.title("읽기 실패 — 상세 오류")
            ew.configure(bg=C["surface"]); ew.geometry("720x480")
            ew.lift(); ew.focus_force()
            tk.Label(ew, text="읽기 실패 — 아래 내용을 스크린샷/복사 후 공유",
                     bg=C["surface"], fg=C["red"],
                     font=(fn,10,"bold")).pack(padx=12, pady=(10,4))
            from tkinter import scrolledtext as _sct2
            box=_sct2.ScrolledText(ew, bg=C["card2"], fg=C["text2"],
                                   font=("Consolas" if IS_WIN else "Courier",9),
                                   relief="flat")
            box.pack(fill="both", expand=True, padx=12, pady=(0,6))
            box.insert("1.0", err_txt); box.config(state="disabled")
            bf3=tk.Frame(ew, bg=C["surface"]); bf3.pack(pady=(0,10))
            ttk.Button(bf3,text="닫기",command=ew.destroy).pack(side="left",padx=6)
            def _cp():
                try:
                    ew.clipboard_clear(); ew.clipboard_append(err_txt); ew.update()
                    messagebox.showinfo("복사됨","클립보드에 복사됐습니다.",parent=ew)
                except Exception: pass
            ttk.Button(bf3,text=" 복사",style="Flat.TButton",
                       command=_cp).pack(side="left",padx=6)

        # ── 라디오 버튼 생성 (함수 정의 완료 후) ─────────────────
        for idx_b, bk in enumerate(books):
            key = bk["name"] if bk["name"] not in wb_map else f"{bk['name']}_{idx_b}"
            bk["_key"] = key; wb_map[key] = bk
            short = bk["path"][-55:] if len(bk["path"])>55 else bk["path"]
            tk.Radiobutton(li,
                text=f"  {bk['name']}  ({len(bk['sheets'])}시트)   …{short}",
                variable=sel_key, value=key,
                bg=C["card"], fg=C["text2"], selectcolor=C["accent"],
                activebackground=C["card"], font=(fn,9),
                command=lambda b=bk: _on_pick(b)
                ).pack(anchor="w", padx=8, pady=3)

        # 첫 번째 자동 선택
        if books:
            _on_pick(books[0])

    def _on_xw_loaded(self, df, wb_name, sh_name, fpath, hdr):
        """열린 엑셀에서 읽은 DataFrame 을 앱에 적용"""
        df = self._clean(df)
        self.raw_xl=None; self.raw_ext=".xlsx"; self.raw_path=fpath
        self.wb_data={sh_name: df}
        self.cmb_sheet["values"]=[sh_name]; self.cmb_sheet.set(sh_name)
        self.spn.delete(0,"end"); self.spn.insert(0,str(hdr))
        self.lbl_hp.config(text="")
        self._apply(df)
        r,c=len(df),len(df.columns)
        self.lbl_file.config(
            text=f"  [열린엑셀]  {wb_name}  /  {sh_name}  ({r:,}행 × {c}열)")
        self._prog(100,f"로드 완료: {wb_name} / {sh_name}")
        self._st(f"열린 엑셀 로드 완료: {wb_name}",C["green"])


    # ======================== FILE LOAD ========================
    def open_file(self,path=None):
        if not path:
            path=filedialog.askopenfilename(
                title="Excel / CSV 파일 선택",
                filetypes=[("Excel/CSV","*.xlsx *.xls *.xlsm *.csv"),("전체","*.*")])
        if not path: return
        self.raw_path=path
        self._st("로딩 중...",C["orange"])
        self._loader.load(path,self.cmb_enc.get())

    def _load_done(self,xl,data,path,ext):
        def _a():
            self.raw_xl=xl; self.raw_ext=ext; self.wb_data=data
            sheets=list(data.keys())
            self.cmb_sheet["values"]=sheets; self.cmb_sheet.set(sheets[0])
            self.spn.delete(0,"end"); self.spn.insert(0,"0"); self.lbl_hp.config(text="")
            self._load_sheet()
            nm=os.path.basename(path); sz=os.path.getsize(path)/1024
            self.lbl_file.config(text=f"  {nm}  ({sz:.0f} KB)  {len(sheets)} sheets")
            _save_recent(path,self._recent); self._recent=_load_recent()
            self._prog(100,f"로드 완료: {nm}")
            self._st(f"완료: {nm}",C["green"])
        self.after(0,_a)

    def _show_recent(self):
        if not self._recent: messagebox.showinfo("최근 파일","최근 파일이 없습니다."); return
        win=tk.Toplevel(self); win.title("최근 파일")
        win.configure(bg=C["surface"]); win.geometry("500x280")
        win.transient(self); win.grab_set()
        tk.Label(win,text="최근 파일",bg=C["surface"],fg=C["text"],
                 font=(FN,11,"bold")).pack(pady=10)
        for p in self._recent:
            def _op(x=p): win.destroy(); self.open_file(x)
            tk.Button(win,text=f"  {os.path.basename(p)}  |  {p}",
                      bg=C["card"],fg=C["text2"],
                      font=(FN,9),
                      relief="flat",anchor="w",cursor="hand2",command=_op
                      ).pack(fill="x",padx=16,pady=2)

    # ======================== SHEET/HEADER ========================
    def _sheet_chg(self):
        self.spn.delete(0,"end"); self.spn.insert(0,"0")
        self.lbl_hp.config(text=""); self._load_sheet()

    def _load_sheet(self):
        nm=self.cmb_sheet.get()
        if nm not in self.wb_data: return
        df=self._clean(self.wb_data[nm].copy()); self._apply(df)

    def _apply_hdr(self):
        try: h=int(self.spn.get())
        except: messagebox.showerror("오류","숫자를 입력하세요"); return
        sh=self.cmb_sheet.get()
        try:
            df=(self.raw_xl.parse(sh,header=h) if self.raw_xl
                else Loader._csv(self.raw_path,self.cmb_enc.get(),h))
            df=self._clean(df); self.wb_data[sh]=df; self._apply(df)
            self._st(f"헤더 {h}행 적용 완료",C["green"])
        except Exception as e: messagebox.showerror("오류",str(e))

    def _ai_hdr(self):
        if not self.raw_xl: return
        sh=self.cmb_sheet.get(); row=HDet.detect(self.raw_xl,sh)
        self.spn.delete(0,"end"); self.spn.insert(0,str(row))
        self._prev_hdr(); self._apply_hdr()
        self._st(f"AI 헤더 자동감지: {row}행",C["cyan"])

    def _prev_hdr(self):
        if not self.raw_xl: return
        try:
            h=int(self.spn.get()); sh=self.cmb_sheet.get()
            pk=self.raw_xl.parse(sh,header=h,nrows=0)
            cs=[str(c).strip() for c in pk.columns if str(c) not in("nan","None","")][:6]
            txt=" | ".join(cs)+("..." if len(pk.columns)>6 else "")
            self.lbl_hp.config(text=f"-> {txt}")
        except: pass

    def _clean(self,df):
        seen={}; new=[]
        for i,c in enumerate(df.columns):
            c=str(c).strip()
            if c in("nan","None",""): c=f"_C{i}"
            orig=c; cnt=seen.get(orig,0)
            if cnt: c=f"{orig}_{cnt}"
            seen[orig]=cnt+1; new.append(c)
        df.columns=new; return df

    def _apply(self,df,undo_label="데이터 변경"):
        self._push_undo(undo_label)
        self.df_raw=df; self.df_view=None
        self.all_cols=list(df.columns); self.vis_cols=list(df.columns)
        self.var_qs.set(""); self.lbl_fr.config(text="")
        self._render_cols()
        self._init_conds(); self._render_tbl(df)
        self._pop_viz(); self._upd_info()
        # populate inspect listboxes
        for lb in(self.lb_icols,self.lb_mcols):
            lb.delete(0,"end")
            for c in self.all_cols: lb.insert("end",c)
        # populate replace col
        self.cmb_rcol["values"]=self.all_cols
        if self.all_cols: self.cmb_rcol.set(self.all_cols[0])

    def _upd_info(self):
        df=self.df_view if self.df_view is not None else self.df_raw
        if df is None: return
        r,c=len(df),len(self.vis_cols)
        tot=len(self.df_raw) if self.df_raw is not None else r
        mem=df.memory_usage(deep=True).sum()/1024/1024
        self.lbl_info.config(text=f"{r:,}/{tot:,} rows\n{c}/{len(self.all_cols)} cols\n{mem:.1f} MB")
        ts = datetime.now().strftime('%H:%M:%S'); self.lbl_st2.config(text=f"{ts}  {r:,}r x {c}c")

    # ======================== COLUMNS ========================
    def _render_cols(self):
        for w in self.col_frame.winfo_children(): w.destroy()
        self.col_vars={}; fn=FN
        for c in self.all_cols:
            v=tk.BooleanVar(value=True); self.col_vars[c]=v
            tk.Checkbutton(self.col_frame,text=c[:30],variable=v,
                bg=C["surface"],fg=C["muted"],selectcolor=C["accent"],
                activebackground=C["surface"],activeforeground=C["text"],
                font=(fn,9),command=self._col_chg,anchor="w").pack(fill="x")

    def _col_chg(self):
        self.vis_cols=[c for c,v in self.col_vars.items() if v.get()]
        self._render_tbl(self.df_view if self.df_view is not None else self.df_raw)
        self._upd_info()

    def _col_all(self,on):
        for v in self.col_vars.values(): v.set(on)
        self.vis_cols=list(self.all_cols) if on else []
        self._render_tbl(self.df_view if self.df_view is not None else self.df_raw)
        self._upd_info()

    # ======================== FILTER ========================
    def _init_conds(self):
        for _,_,_,fr in self.cond_rows: fr.destroy()
        self.cond_rows=[]; self._add_cond()

    def _add_cond(self):
        fr=tk.Frame(self.cond_frame,bg=C["card"]); fr.pack(fill="x",pady=2)
        cv=tk.StringVar(value=self.all_cols[0] if self.all_cols else "")
        vv=tk.StringVar()
        fn=FN
        cc=ttk.Combobox(fr,textvariable=cv,values=self.all_cols,width=18,state="readonly")
        cc.pack(side="left",padx=(0,4))
        co=ttk.Combobox(fr,width=11,state="readonly",values=[l for l,_ in OPS])
        co.set("포함"); co.pack(side="left",padx=(0,4))
        ent=ttk.Entry(fr,textvariable=vv,width=24); ent.pack(side="left",padx=(0,4))
        ent.bind("<Return>",lambda e:self._do_filter())
        tk.Label(fr,text="AND",bg=C["card"],fg=C["muted2"],font=(fn,8)).pack(side="left",padx=4)
        def _d(f=fr):
            f.destroy(); self.cond_rows=[(a,b,c,d) for a,b,c,d in self.cond_rows if d!=f]
        tk.Button(fr,text="X",bg="#2d1515",fg=C["red"],bd=0,
                  font=(fn,9,"bold"),cursor="hand2",command=_d).pack(side="left")
        self.cond_rows.append((cv,co,vv,fr))

    def _opk(self,cmb):
        l=cmb.get()
        for lb,k in OPS:
            if lb==l: return k
        return "ct"

    def _do_filter(self):
        if self.df_raw is None: return
        self._push_undo("필터 적용")
        conds=[(cv.get(),self._opk(co),vv.get()) for cv,co,vv,_ in self.cond_rows if cv.get()]
        res=FEng.run(self.df_raw,conds); self.df_view=res
        self._render_tbl(res)
        pct=len(res)/max(len(self.df_raw),1)*100
        self.lbl_fr.config(text=f"{len(res):,} rows ({pct:.1f}%)")
        self._upd_info(); self._st(f"필터 결과: {len(res):,}행",C["cyan"])

    def _reset_filter(self):
        self.df_view=None; self.var_qs.set(""); self.lbl_fr.config(text="")
        self._render_tbl(self.df_raw); self._upd_info(); self._st("초기화",C["muted"])

    def _qs(self):
        if self._qjob: self.after_cancel(self._qjob)
        self._qjob=self.after(300,self._qs_run)

    def _qs_run(self):
        q=self.var_qs.get().strip()
        if self.df_raw is None: return
        if not q:
            self.df_view=None; self._render_tbl(self.df_raw)
            self.lbl_fr.config(text=""); self._upd_info(); return
        mask=self.df_raw.apply(
            lambda col:col.astype(str).str.contains(q,case=False,na=False)).any(axis=1)
        self.df_view=self.df_raw[mask]; self._render_tbl(self.df_view)
        self.lbl_fr.config(text=f"검색 결과: {len(self.df_view):,}행"); self._upd_info()

    # ======================== TABLE ========================
    def _render_tbl(self,df):
        if df is None: return
        self.tree.load(df,self.vis_cols)
        shown=min(VTree.PAGE,len(df))
        self.lbl_tbl.config(text=f"Data  ({shown:,} / {len(df):,} rows)")
        self._upd_pg()

    def _upd_pg(self):
        if not self.tree._vd: return
        o=self.tree.offset; e=min(o+VTree.PAGE,self.tree.total)
        self.lbl_pg.config(text=f"[{o+1}-{e} / {self.tree.total:,}]")

    def _pnext(self):
        o=self.tree.offset+VTree.PAGE
        if o<self.tree.total: self.tree._render(o); self._upd_pg()

    def _pprev(self):
        o=max(0,self.tree.offset-VTree.PAGE); self.tree._render(o); self._upd_pg()

    def _copy(self,e):
        sel=self.tree.selection()
        if not sel: return
        rows=["	".join(str(v) for v in self.tree.item(s,"values")) for s in sel]
        self.clipboard_clear(); self.clipboard_append("\n".join(rows))
        self._st(f"{len(rows)}행 복사됨",C["cyan"])

    # ======================== MERGE ========================
    def _load_merge_file(self):
        p=filedialog.askopenfilename(
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not p: return
        try:
            ext=os.path.splitext(p)[1].lower()
            if ext in(".xlsx",".xlsm"): df=pd.ExcelFile(p,engine="openpyxl").parse(0)
            elif ext==".xls":           df=pd.ExcelFile(p,engine="xlrd").parse(0)
            else:                       df=Loader._csv(p,self.cmb_enc.get(),0)
            df=self._clean(df); self._ext_df=df
            self.lbl_mf.config(text=os.path.basename(p))
            self.cmb_ek["values"]=list(df.columns)
            if df.columns.any(): self.cmb_ek.set(df.columns[0])
            self.lb_mcols.delete(0,"end")
            for c in df.columns: self.lb_mcols.insert("end",c)
            self._st(f"외부 파일: {os.path.basename(p)}  ({len(df):,}행)",C["teal"])
        except Exception as e: messagebox.showerror("오류",str(e))

    def _do_merge(self):
        if self.df_raw is None or self._ext_df is None:
            messagebox.showwarning("주의","메인 파일과 외부 파일을 모두 불러오세요."); return
        mk=self.cmb_mk.get(); ek=self.cmb_ek.get()
        jt=self.cmb_join.get()
        sel=self.lb_mcols.curselection()
        if not sel: messagebox.showwarning("주의","추가할 컬럼을 선택하세요."); return
        ecols=[self.lb_mcols.get(i) for i in sel]
        if ek not in ecols: ecols=[ek]+ecols
        try:
            src=self.df_view if self.df_view is not None else self.df_raw
            right=self._ext_df[ecols].copy()
            merged=src.merge(right,left_on=mk,right_on=ek,how=jt,suffixes=("","_ext"))
            self._merge_df=merged
            self.mtree.load(merged,list(merged.columns))
            matched=merged[ek+"_ext"].notna().sum() if ek+"_ext" in merged.columns else len(merged)
            self.lbl_mr.config(
                text=f"병합 완료: {len(merged):,}행  |  매칭: {matched:,}")
            self._st(f"병합 완료: {len(merged):,}행",C["green"])
        except Exception as e: messagebox.showerror("병합 오류",str(e))

    def _save_merged(self):
        if self._merge_df is None:
            messagebox.showinfo("알림","먼저 병합을 실행하세요."); return
        self._save_df(self._merge_df)

    # ======================== REPLACE ========================
    def _add_rep_row(self):
        fr=tk.Frame(self.rep_frame,bg=C["card"]); fr.pack(fill="x",pady=1)
        fv=tk.StringVar(); tv=tk.StringVar()
        ttk.Entry(fr,textvariable=fv,width=24).pack(side="left",padx=(0,4))
        ttk.Entry(fr,textvariable=tv,width=24).pack(side="left",padx=(0,4))
        def _d(f=fr,row=(fv,tv,fr)):
            f.destroy(); self.rep_rows=[r for r in self.rep_rows if r!=row]
        tk.Button(fr,text="X",bg="#2d1515",fg=C["red"],bd=0,
                  font=(FN,9,"bold"),
                  cursor="hand2",command=_d).pack(side="left")
        self.rep_rows.append((fv,tv,fr))

    def _load_rep_csv(self):
        p=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")])
        if not p: return
        try:
            df=Loader._csv(p,"auto",0)
            if len(df.columns)<2: raise ValueError("컬럼이 2개 이상 필요합니다 (원본, 변경)")
            for _,row in df.iterrows(): self._add_rep_row()
            # fill last N rows
            pairs=list(zip(df.iloc[:,0].astype(str),df.iloc[:,1].astype(str)))
            tail=self.rep_rows[-len(pairs):]
            for (fv,tv,_),(_f,_t) in zip(tail,pairs):
                fv.set(_f); tv.set(_t)
            self._st(f"치환 규칙 {len(pairs)}개 로드 완료",C["teal"])
        except Exception as e: messagebox.showerror("오류",str(e))

    def _do_replace(self):
        if self.df_raw is None: return
        col=self.cmb_rcol.get(); mode=self.cmb_rmode.get()
        rules=[(fv.get(),tv.get()) for fv,tv,_ in self.rep_rows
               if fv.get().strip()!=""]
        if not rules: messagebox.showwarning("주의","치환 규칙을 입력하세요."); return
        df=(self.df_view if self.df_view is not None else self.df_raw).copy()
        if col not in df.columns:
            messagebox.showerror("Error",f"컬럼 '{col}'을 찾을 수 없습니다."); return
        changed=0
        for frm,to in rules:
            if mode=="완전일치":
                mask=df[col].astype(str)==frm
                changed+=mask.sum(); df.loc[mask,col]=to
            elif mode=="포함":
                mask=df[col].astype(str).str.contains(re.escape(frm),na=False)
                changed+=mask.sum()
                df.loc[mask,col]=df.loc[mask,col].astype(str).str.replace(
                    re.escape(frm),to,regex=False)
            elif mode=="정규식":
                mask=df[col].astype(str).str.contains(frm,regex=True,na=False)
                changed+=mask.sum()
                df.loc[mask,col]=df.loc[mask,col].astype(str).str.replace(
                    frm,to,regex=True)
        self.df_view=df; self.df_raw=df
        self._render_tbl(df); self.rtree.load(df,self.vis_cols)
        self.lbl_rr.config(text=f"치환 완료: {changed:,}셀"); self._upd_info()
        self._st(f"치환: {changed:,}셀 변경",C["green"])

    def _reset_replace(self):
        self._load_sheet()
        self.lbl_rr.config(text="원본으로 초기화"); self._st("초기화",C["muted"])

    # ======================== INSPECT ========================
    def _do_inspect(self):
        if self.df_raw is None: return
        df=self.df_view if self.df_view is not None else self.df_raw
        sel=self.lb_icols.curselection()
        key_cols=[self.lb_icols.get(i) for i in sel] if sel else list(df.columns)

        issues=[]
        # duplicates
        dup_mask=df.duplicated(subset=key_cols,keep=False)
        dup_df=df[dup_mask].copy()
        dup_df["__ISSUE__"]="중복"
        issues.append(dup_df)
        # blanks
        for c in df.columns:
            blank_mask=df[c].astype(str).str.strip().isin(["","nan","None"])
            bl=df[blank_mask].copy()
            bl["__ISSUE__"]=f"공백:{c}"
            issues.append(bl)

        if issues:
            all_issues=pd.concat(issues,ignore_index=True).drop_duplicates()
        else:
            all_issues=pd.DataFrame(columns=list(df.columns)+["__ISSUE__"])

        self._issues_df=all_issues
        n_dup=len(dup_df); n_bl=sum(len(i) for i in issues[1:])

        # KPI
        for w in self.kpi_frame.winfo_children(): w.destroy()
        self._kpi_cards(self.kpi_frame,[
            ("전체 행",f"{len(df):,}",C["text"]),
            ("중복 행",f"{n_dup:,}",C["red"] if n_dup else C["green"]),
            ("공백 셀",f"{n_bl:,}",C["yellow"] if n_bl else C["green"]),
            ("총 이슈",f"{len(all_issues):,}",C["orange"] if all_issues.__len__() else C["green"]),
        ])
        self.itree.load(all_issues,list(all_issues.columns))
        self.lbl_ir.config(
            text=f"중복: {n_dup:,}  공백: {n_bl:,}  총 이슈: {len(all_issues):,}")
        self._st(f"점검 완료: {len(all_issues):,}건",
                 C["red"] if len(all_issues) else C["green"])

    def _exp_issues(self):
        if self._issues_df is None:
            messagebox.showinfo("알림","먼저 점검을 실행하세요."); return
        self._save_df(self._issues_df)

    def _kpi_cards(self,parent,items):
        for lbl,val,color in items:
            fr=tk.Frame(parent,bg=C["card"],padx=14,pady=8)
            fr.pack(side="left",fill="x",expand=True,padx=4,pady=4)
            fn=FN
            tk.Label(fr,text=lbl,bg=C["card"],fg=C["muted"],font=(fn,8)).pack(anchor="w")
            tk.Label(fr,text=val, bg=C["card"],fg=color, font=(fn,14,"bold")).pack(anchor="w")

    # ======================== CHART / KPI ========================
    def _pop_viz(self):
        for a in("cmb_vg","cmb_vv","cmb_pr","cmb_pv","cmb_pa"):
            getattr(self,a)["values"]=self.all_cols
            if self.all_cols: getattr(self,a).set(self.all_cols[0])
        if len(self.all_cols)>1:
            self.cmb_vv.set(self.all_cols[1])
            self.cmb_pv.set(self.all_cols[1])
        self.cmb_pc["values"]=["(없음)"]+self.all_cols; self.cmb_pc.set("(없음)")
        self.cmb_mk["values"]=self.all_cols
        if self.all_cols: self.cmb_mk.set(self.all_cols[0])

    def _src(self):
        return self.df_view if self.df_view is not None else self.df_raw

    def _auto_kpi(self):
        if self.df_raw is None: return
        find=lambda *ks:next((c for c in self.all_cols if any(k in c for k in ks)),None)
        grp=find("관리지사","지사","branch","region","지점","부서","본부","센터","팀")
        fee=find("월정료","월정","fee","amount","금액","요금","charge","단가","월액")
        if not grp:
            messagebox.showwarning("주의","'관리지사' 컬럼을 찾을 수 없습니다."); return
        src=self._src()
        kw={"계약건수":(grp,"count")}
        if fee: kw["월정료합계"]=(fee,"sum")
        agg=src.groupby(grp).agg(**kw).reset_index()
        sc="월정료합계" if "월정료합계" in agg.columns else "계약건수"
        agg=agg.sort_values(sc,ascending=False)

        for w in self.kpi_vframe.winfo_children(): w.destroy()
        kpis=[("지사 수",f"{len(agg):,}",C["accent"]),
              ("계약 건수",f"{agg['계약건수'].sum():,}",C["green"])]
        if "월정료합계" in agg.columns:
            kpis.append(("월정료 합계",f"{agg['월정료합계'].sum():,.0f}",C["purple"]))
        self._kpi_cards(self.kpi_vframe,kpis)

        ct=self.cmb_vt.get(); lbl=agg[grp].astype(str).tolist()
        cols=(CHART*5)[:len(lbl)]
        if ct in("파이","도넛"):
            vals=agg.get("월정료합계",agg["계약건수"]).values
            fig,ax=plt.subplots(figsize=(8,5),facecolor=C["card"])
            ax.set_facecolor(C["card"]); CEng.pie(ax,lbl,vals,cols,donut=ct=="도넛")
            ax.set_title("지사별 분포",color=C["text"],fontsize=11)
            fig.tight_layout(); self._show(fig)
        else:
            nc=2 if "월정료합계" in agg.columns else 1
            fig,axes=plt.subplots(1,nc,figsize=(11,5),facecolor=C["card"])
            if nc==1: axes=[axes]
            h=ct=="가로막대"
            for ax in axes: ax.set_facecolor(C["card"])
            axes[0].set_title("계약 건수",color=C["text"],fontsize=10)
            CEng.bar(axes[0],lbl,agg["계약건수"].values,cols,horiz=h)
            if nc==2:
                axes[1].set_title("월정료 합계",color=C["text"],fontsize=10)
                CEng.bar(axes[1],lbl,agg["월정료합계"].values,[CHART[1]]*len(lbl),horiz=h)
            fig.tight_layout(); self._show(fig)
        self.nb.select(4); self._st(f"KPI: {len(agg)}개 지사",C["purple"])

    def _build_chart(self):
        if self.df_raw is None: return
        g=self.cmb_vg.get(); v=self.cmb_vv.get()
        m=self.cmb_va.get(); ct=self.cmb_vt.get()
        if not g or not v: return
        src=self._src().copy(); fn=AGG.get(m,"sum")
        if fn=="count":
            agg=src.groupby(g)[v].count().reset_index()
        else:
            src[v]=pd.to_numeric(src[v],errors="coerce")
            agg=src.groupby(g)[v].agg(fn).reset_index()
        agg.columns=[g,"val"]
        agg=agg.sort_values("val",ascending=False).head(25)
        lbl=agg[g].astype(str).tolist(); vals=agg["val"].values
        cols=(CHART*5)[:len(lbl)]

        for w in self.kpi_vframe.winfo_children(): w.destroy()
        self._kpi_cards(self.kpi_vframe,[
            ("항목 수",f"{len(agg):,}",C["text"]),
            ("합계",f"{vals.sum():,.1f}",C["accent"]),
            ("평균", f"{vals.mean():,.1f}",C["green"]),
            ("최대",  f"{vals.max():,.1f}",C["orange"]),
        ])
        fig,ax=plt.subplots(figsize=(10,5),facecolor=C["card"])
        ax.set_facecolor(C["card"])
        if   ct=="파이":    CEng.pie(ax,lbl,vals,cols)
        elif ct=="도넛":  CEng.pie(ax,lbl,vals,cols,donut=True)
        elif ct=="가로막대":  CEng.bar(ax,lbl,vals,cols,horiz=True)
        elif ct=="꺾은선":   CEng.line(ax,lbl,vals,CHART[0])
        elif ct=="누적막대":
            CEng.ax(ax); bot=np.zeros(1)
            for i,(l2,v2) in enumerate(zip(lbl,vals)):
                ax.bar([g],[v2],bottom=bot,color=CHART[i%len(CHART)],label=l2[:12]); bot+=v2
            ax.legend(loc="upper right",fontsize=7,framealpha=0.3)
        else:
            CEng.bar(ax,lbl,vals,cols)
        ax.set_title(f"{g} x {v} ({m})",fontsize=11,color=C["text"])
        fig.tight_layout(); self._show(fig); self.nb.select(4)

    def _show(self,fig):
        plt.close("all")
        for w in self.chart_frame.winfo_children(): w.destroy()
        fig.patch.set_facecolor(C["card"])
        cv=FigureCanvasTkAgg(fig,master=self.chart_frame)
        cv.draw(); cv.get_tk_widget().pack(fill="both",expand=True)
        self._cur_fig=fig

    def _clr_chart(self):
        plt.close("all")
        for w in self.chart_frame.winfo_children(): w.destroy()
        fn=FN
        tk.Label(self.chart_frame,text="  차트를 생성하면 여기에 표시됩니다\n\n상단의 [그룹(X)] [집계값(Y)] 설정 후 [차트 생성] 클릭",
                 bg=C["card"],fg=C["muted2"],font=(fn,12)).pack(expand=True)
        self._cur_fig=None
        for w in self.kpi_vframe.winfo_children(): w.destroy()

    def _save_chart(self):
        if not self._cur_fig:
            messagebox.showinfo("알림","먼저 차트를 생성하세요."); return
        p=filedialog.asksaveasfilename(defaultextension=".png",
            filetypes=[("PNG 이미지","*.png"),("SVG","*.svg"),("PDF","*.pdf")])
        if not p: return
        self._cur_fig.savefig(p,dpi=180,bbox_inches="tight",facecolor=C["card"])
        self._st(f"차트 저장: {os.path.basename(p)}",C["green"])

    # ======================== PIVOT ========================
    def _pivot(self):
        if self.df_raw is None: return
        rc=self.cmb_pr.get(); cc2=self.cmb_pc.get(); vc=self.cmb_pv.get()
        m=self.cmb_pa.get(); fn=AGG.get(m,"sum")
        src=self._src().copy()
        src[vc]=pd.to_numeric(src[vc],errors="coerce")
        try:
            if cc2 and cc2!="(없음)":
                pv=src.pivot_table(index=rc,columns=cc2,values=vc,
                                   aggfunc=fn,fill_value=0,
                                   margins=True,margins_name="[합계]").reset_index()
                pv.columns=[str(c) for c in pv.columns]
            else:
                if fn=="count": pv=src.groupby(rc)[vc].count().reset_index()
                else:           pv=src.groupby(rc)[vc].agg(fn).reset_index()
                pv.columns=[rc,f"{vc}_{m}"]
                pv=pv.sort_values(f"{vc}_{m}",ascending=False)
                tot={rc:"[합계]",f"{vc}_{m}":pv[f"{vc}_{m}"].sum()}
                pv=pd.concat([pv,pd.DataFrame([tot])],ignore_index=True)
        except Exception as e:
            messagebox.showerror("피벗 오류",str(e)); return
        self._pivot_df=pv; cs=list(pv.columns)
        self.ptree["columns"]=cs
        for c in cs:
            self.ptree.heading(c,text=c)
            self.ptree.column(c,width=max(100,len(c)*10),anchor="w")
        for it in self.ptree.get_children(): self.ptree.delete(it)
        for i,(_,row) in enumerate(pv.iterrows()):
            vs=[]
            for v in row:
                if pd.isna(v): vs.append("")
                elif isinstance(v,(int,np.integer)): vs.append(f"{v:,}")
                elif isinstance(v,float): vs.append(f"{v:,.2f}")
                else: vs.append(str(v))
            tg="total" if str(row.iloc[0]).startswith("[합계") else("odd" if i%2 else "")
            self.ptree.insert("","end",values=vs,tags=(tg,))
        n=len(pv)-1
        self.lbl_pv.config(text=f"집계 결과  {n:,}항목")
        self._st(f"피벗 완료: {n:,}행",C["green"])

    def _exp_pivot(self):
        if self._pivot_df is None:
            messagebox.showinfo("알림","먼저 집계를 실행하세요."); return
        self._save_df(self._pivot_df)

    def _pivot_chart(self):
        if self._pivot_df is None:
            messagebox.showinfo("알림","먼저 집계를 실행하세요."); return
        cs=list(self._pivot_df.columns)
        if len(cs)<2: return
        g=cs[0]; v2=cs[1] if len(cs)==2 else cs[-1]
        lbl=self._pivot_df[g].astype(str).tolist()
        vals=pd.to_numeric(self._pivot_df[v2],errors="coerce").fillna(0).values
        cols=(CHART*5)[:len(lbl)]
        fig,ax=plt.subplots(figsize=(10,5),facecolor=C["card"])
        ax.set_facecolor(C["card"]); CEng.bar(ax,lbl,vals,cols)
        ax.set_title(f"{g} x {v2}",fontsize=11,color=C["text"])
        fig.tight_layout(); self._show(fig); self.nb.select(4)

    # ======================== STATS ========================
    def _stats(self):
        if self.df_raw is None: return
        src=self._src(); nc=src.select_dtypes(include="number").columns.tolist()
        if not nc: messagebox.showinfo("알림","수치형 컬럼이 없습니다."); return
        rows=[]
        for c in nc:
            s=src[c].dropna()
            if not len(s): continue
            rows.append({"컬럼":c,"건수":f"{len(s):,}","결측":f"{src[c].isna().sum():,}",
                         "합계":f"{s.sum():,.2f}","평균":f"{s.mean():,.2f}",
                         "중앙값":f"{s.median():,.2f}","표준편차":f"{s.std():,.2f}",
                         "최솟값":f"{s.min():,.2f}","최대":f"{s.max():,.2f}",
                         "25%":f"{s.quantile(.25):,.2f}","75%":f"{s.quantile(.75):,.2f}"})
        if not rows: return
        df2=pd.DataFrame(rows); self._stats_df=df2; cs=list(df2.columns)
        self.stree["columns"]=cs
        for c in cs:
            self.stree.heading(c,text=c)
            self.stree.column(c,width=max(70,len(c)*9),anchor="e" if c!="컬럼" else "w")
        for it in self.stree.get_children(): self.stree.delete(it)
        for i,(_,row) in enumerate(df2.iterrows()):
            tg="odd" if i%2 else ""
            self.stree.insert("","end",values=list(row),tags=(tg,))
        self._st(f"통계 완료: 수치 컬럼 {len(rows)}개",C["green"])

    def _exp_stats(self):
        if self._stats_df is None:
            messagebox.showinfo("알림","먼저 통계를 계산하세요."); return
        self._save_df(self._stats_df)

    # ======================== EXPORT ========================
    def _exp(self,fmt):
        if self.df_raw is None: return
        src=self.df_view if self.df_view is not None else self.df_raw
        out=src[[c for c in self.vis_cols if c in src.columns]]
        self._save_df(out,fmt)

    def _exp_filt(self):
        if self.df_view is None:
            messagebox.showinfo("알림","먼저 필터를 적용하세요."); return
        out=self.df_view[[c for c in self.vis_cols if c in self.df_view.columns]]
        self._save_df(out)

    def _save_df(self,df,fmt=None):
        if fmt is None:
            p=filedialog.asksaveasfilename(defaultextension=".xlsx",
                filetypes=[("Excel","*.xlsx"),("CSV","*.csv"),("전체","*.*")])
        else:
            p=filedialog.asksaveasfilename(defaultextension=f".{fmt}",
                filetypes=[(fmt.upper(),f"*.{fmt}"),("All","*.*")])
        if not p: return
        try:
            if p.lower().endswith(".csv"):
                df.to_csv(p,index=False,encoding="utf-8-sig")
            else:
                with pd.ExcelWriter(p,engine="openpyxl") as w:
                    df.to_excel(w,index=False,sheet_name="결과")
            self._st(f"저장 완료: {os.path.basename(p)} ({len(df):,}행)",C["green"])
            messagebox.showinfo("저장 완료",f"{p}\n({len(df):,}행 저장됨)")
        except Exception as e: messagebox.showerror("저장 오류",str(e))


    # ─────────────────────────────────────────────────────────────
    # TAB 8: 위경도 변환 (카카오 API)
    # ─────────────────────────────────────────────────────────────
    def _t_geocode(self):
        import re, time, random, threading
        try: import requests, urllib3
        except ImportError:
            import subprocess,sys
            subprocess.check_call([sys.executable,"-m","pip","install","requests","urllib3","-q"],
                                  stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            import requests, urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        tab=ttk.Frame(self.nb); self.nb.add(tab,text="   위경도 변환  ")

        # ── 설정 카드 ──────────────────────────────────────────────
        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="카카오 API  주소 → 위경도 변환",
                 bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")
                 ).pack(anchor="w",padx=12,pady=(8,4))

        # API 키 + 주소 컬럼
        r1=tk.Frame(cc,bg=C["card"]); r1.pack(fill="x",padx=12,pady=2)
        tk.Label(r1,text="카카오 REST API 키:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.geo_api_key=tk.StringVar(value="af04a0a8e5416c95eaa04cccc060031d")
        ttk.Entry(r1,textvariable=self.geo_api_key,width=38,show="").pack(side="left",padx=6)

        r2=tk.Frame(cc,bg=C["card"]); r2.pack(fill="x",padx=12,pady=2)
        tk.Label(r2,text="주소 컬럼명:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        self.geo_addr_col=ttk.Combobox(r2,state="normal",width=20)
        self.geo_addr_col.set("설치주소")
        self.geo_addr_col.pack(side="left",padx=6)
        tk.Label(r2,text="헤더 행:",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left",padx=(12,0))
        self.geo_hdr_row=tk.Spinbox(r2,from_=1,to=10,width=4,
            bg=C["card2"],fg=C["text"],insertbackground=C["text"],
            buttonbackground=C["border"],relief="flat")
        self.geo_hdr_row.delete(0,"end"); self.geo_hdr_row.insert(0,"1")
        self.geo_hdr_row.pack(side="left",padx=4)

        # 옵션
        r3=tk.Frame(cc,bg=C["card"]); r3.pack(fill="x",padx=12,pady=4)
        self.geo_skip_exist=tk.BooleanVar(value=True)
        tk.Checkbutton(r3,text="이미 좌표 있는 행 건너뛰기",variable=self.geo_skip_exist,
                       bg=C["card"],fg=C["text"],selectcolor=C["accent"],
                       activebackground=C["card"],font=(FN,9)
                       ).pack(side="left")
        self.geo_delay=tk.DoubleVar(value=0.15)
        tk.Label(r3,text="  딜레이(초):",bg=C["card"],fg=C["muted"],
                 font=(FN,9)).pack(side="left",padx=(16,0))
        ttk.Entry(r3,textvariable=self.geo_delay,width=5).pack(side="left",padx=4)

        # 실행 버튼 행
        bf=tk.Frame(cc,bg=C["card"]); bf.pack(fill="x",padx=12,pady=(4,10))
        self.geo_btn=ttk.Button(bf,text=" 위경도 변환 시작",
                                command=self._geo_run)
        self.geo_btn.pack(side="left")
        self.geo_stop_flag=tk.BooleanVar(value=False)
        ttk.Button(bf,text="⏹ 중지",style="Flat.TButton",
                   command=lambda:self.geo_stop_flag.set(True)).pack(side="left",padx=6)
        ttk.Button(bf,text=" 컬럼 새로고침",style="Flat.TButton",
                   command=self._geo_refresh_cols).pack(side="left")

        # KPI 행
        self.geo_kpi=tk.Frame(tab,bg=C["bg"]); self.geo_kpi.pack(fill="x",padx=6,pady=2)

        # 진행 바
        pf=tk.Frame(tab,bg=C["bg"]); pf.pack(fill="x",padx=6,pady=(0,2))
        self.geo_pb_var=tk.DoubleVar(value=0)
        self.geo_pb=ttk.Progressbar(pf,variable=self.geo_pb_var,
                                    style="Horizontal.TProgressbar",maximum=100)
        self.geo_pb.pack(fill="x",expand=True)
        self.geo_lbl_prog=tk.Label(pf,text="",bg=C["bg"],fg=C["muted"],
                                   font=(FN,8))
        self.geo_lbl_prog.pack(anchor="e")

        # 로그창
        lf=tk.Frame(tab,bg=C["card"]); lf.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(lf,text="변환 로그",bg=C["card"],fg=C["text"],
                 font=(FN,10,"bold")
                 ).pack(anchor="w",padx=10,pady=(8,4))
        from tkinter import scrolledtext as _st_mod
        self.geo_log=_st_mod.ScrolledText(lf,height=14,bg=C["card2"],fg=C["text2"],
                                          insertbackground=C["text"],
                                          font=("Consolas" if IS_WIN else "Courier",9),
                                          relief="flat")
        self.geo_log.pack(fill="both",expand=True,padx=8,pady=(0,8))

        # 내부 유틸 함수들 (클로저로 보관)
        def _is_blank(x):
            return x is None or str(x).strip() in ("","nan","None")

        def _clean_addr(addr):
            addr=str(addr)
            addr=re.sub(r"\(.*?\)","",addr)
            addr=re.sub(r"\d+층|\d+호","",addr)
            addr=re.sub(r"[^\w\s가-힣\d-]"," ",addr)
            addr=re.sub(r"\s+"," ",addr).strip()
            return addr

        def _split_addr(addr):
            addr=str(addr); 시=군구=읍면동=""
            m1=re.search(r"(서울|부산|대구|인천|광주|대전|울산|세종|제주|경기|강원|충북|충남|전북|전남|경북|경남)",addr)
            m2=re.search(r"([가-힣]+구|[가-힣]+시|[가-힣]+군)",addr)
            m3=re.search(r"([가-힣]+[읍면동])",addr)
            if m1: 시=m1.group(1)
            if m2: 군구=m2.group(1)
            if m3: 읍면동=m3.group(1)
            return 시,군구,읍면동

        def _kakao_geo(session,query,api_key,retry=3):
            hdrs={"Authorization":f"KakaoAK {api_key}"}
            a_url="https://dapi.kakao.com/v2/local/search/address.json"
            k_url="https://dapi.kakao.com/v2/local/search/keyword.json"
            cleaned=_clean_addr(query)
            for _ in range(retry+1):
                try:
                    r=session.get(a_url,headers=hdrs,params={"query":cleaned},timeout=10,verify=False)
                    if r.status_code==429: time.sleep(1.2+random.random()); continue
                    if r.status_code in(401,403): return None,None,"AUTH"
                    if r.status_code!=200: return None,None,"HTTP"
                    docs=r.json().get("documents",[])
                    if docs: return float(docs[0]["y"]),float(docs[0]["x"]),"ADDRESS"
                    r2=session.get(k_url,headers=hdrs,params={"query":cleaned},timeout=10,verify=False)
                    if r2.status_code==429: time.sleep(1.2+random.random()); continue
                    if r2.status_code in(401,403): return None,None,"AUTH"
                    if r2.status_code!=200: return None,None,"HTTP"
                    docs2=r2.json().get("documents",[])
                    if docs2: return float(docs2[0]["y"]),float(docs2[0]["x"]),"KEYWORD"
                    return None,None,"NO_RESULT"
                except Exception: time.sleep(0.5)
            return None,None,"EXCEPTION"

        # 탭 인스턴스에 함수 바인딩
        self._geo_is_blank  = _is_blank
        self._geo_clean     = _clean_addr
        self._geo_split     = _split_addr
        self._geo_kakao     = _kakao_geo

    def _geo_log(self,msg):
        self.geo_log.insert(tk.END,msg+"\n")
        self.geo_log.see(tk.END)
        self.update_idletasks()

    def _geo_kpi_cards(self,items):
        for w in self.geo_kpi.winfo_children(): w.destroy()
        fn=FN
        for lbl,val,color in items:
            fr=tk.Frame(self.geo_kpi,bg=C["card"],padx=12,pady=6)
            fr.pack(side="left",fill="x",expand=True,padx=4,pady=4)
            tk.Label(fr,text=lbl,bg=C["card"],fg=C["muted"],font=(fn,8)).pack(anchor="w")
            tk.Label(fr,text=val, bg=C["card"],fg=color, font=(fn,13,"bold")).pack(anchor="w")

    def _geo_refresh_cols(self):
        """현재 로드된 데이터의 컬럼명을 주소 컬럼 드롭다운에 채움"""
        if self.all_cols:
            self.geo_addr_col["values"]=self.all_cols
            self._geo_log("컬럼 목록 새로고침 완료.")
        else:
            self._geo_log("⚠ 먼저 파일을 불러오세요.")

    def _geo_run(self):
        """위경도 변환 메인 실행 (비동기 스레드)"""
        import threading
        self.geo_stop_flag.set(False)
        self.geo_btn.config(state="disabled")
        self.geo_log.delete(1.0,tk.END)
        threading.Thread(target=self._geo_process,daemon=True).start()

    def _geo_process(self):
        import requests, urllib3, time, random
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        try:
            api_key  = self.geo_api_key.get().strip()
            addr_col = self.geo_addr_col.get().strip()
            try: hdr_row=int(self.geo_hdr_row.get())
            except: hdr_row=1

            if not api_key:
                self.after(0,lambda:messagebox.showerror("오류","카카오 API 키를 입력하세요.")); return
            if self.df_raw is None:
                self.after(0,lambda:messagebox.showwarning("주의","파일을 먼저 불러오세요.")); return
            if addr_col not in self.df_raw.columns:
                self.after(0,lambda:messagebox.showerror("오류",f"'{addr_col}' 컬럼을 찾을 수 없습니다.")); return

            df = (self.df_view if self.df_view is not None else self.df_raw).copy()
            n  = len(df)
            self.after(0,lambda:self._geo_log(f" 주소 데이터 {n:,}건 감지"))
            self.after(0,lambda:self.geo_pb.__setitem__("maximum",100))

            # 결과 컬럼 초기화
            RESULT_COLS=["위도","경도","시","군구","읍면동",
                         "위치좌표(위도,경도)","지도링크_URL"]
            for c in RESULT_COLS:
                if c not in df.columns: df[c]=""

            session=requests.Session()
            ok=fail=skip=0
            delay=max(0.05,self.geo_delay.get())

            for i,(idx,row) in enumerate(df.iterrows()):
                if self.geo_stop_flag.get():
                    self.after(0,lambda:self._geo_log("⏹ 사용자가 중지했습니다.")); break

                pct=int((i+1)/n*100)
                prog_txt=f"처리 중... ({i+1:,}/{n:,})"
                self.after(0,lambda p=pct,t=prog_txt:(
                    self.geo_pb_var.set(p),
                    self.geo_lbl_prog.config(text=t)
                ))

                addr=row[addr_col]
                if self._geo_is_blank(addr):
                    continue

                # 이미 좌표 존재 시 건너뛰기
                if self.geo_skip_exist.get():
                    if not self._geo_is_blank(row.get("위도","")) and                        not self._geo_is_blank(row.get("경도","")):
                        try:
                            lat2=float(row["위도"]); lng2=float(row["경도"])
                            si,gungu,eup=self._geo_split(addr)
                            df.at[idx,"위도"]=lat2; df.at[idx,"경도"]=lng2
                            df.at[idx,"시"]=si; df.at[idx,"군구"]=gungu; df.at[idx,"읍면동"]=eup
                            df.at[idx,"위치좌표(위도,경도)"]=f"{lat2},{lng2}"
                            df.at[idx,"지도링크_URL"]=f"https://www.google.com/maps/search/?api=1&query={lat2},{lng2}"
                            skip+=1
                            msg=f"[{i+1}/{n}] ⏭ 패스(기존좌표): {str(addr)[:40]}"
                            self.after(0,lambda m=msg:self._geo_log(m)); continue
                        except: pass

                # API 호출
                lat,lng,mode=self._geo_kakao(session,addr,api_key,retry=3)
                si,gungu,eup=self._geo_split(addr)
                df.at[idx,"시"]=si; df.at[idx,"군구"]=gungu; df.at[idx,"읍면동"]=eup

                if lat is None:
                    fail+=1
                    reason={"AUTH":"API키오류","NO_RESULT":"검색없음",
                            "HTTP":"HTTP오류","EXCEPTION":"예외"}.get(mode,mode)
                    msg=f"[{i+1}/{n}]  실패({reason}): {str(addr)[:35]}"
                else:
                    ok+=1
                    df.at[idx,"위도"]=lat; df.at[idx,"경도"]=lng
                    df.at[idx,"위치좌표(위도,경도)"]=f"{lat},{lng}"
                    df.at[idx,"지도링크_URL"]=f"https://www.google.com/maps/search/?api=1&query={lat},{lng}"
                    msg=f"[{i+1}/{n}]  완료: {str(addr)[:35]}  →  {lat:.6f}, {lng:.6f}"

                self.after(0,lambda m=msg:self._geo_log(m))

                # KPI 실시간 업데이트
                kpis=[
                    ("전체",   f"{n:,}건",     C["text"]),
                    (" 성공", f"{ok:,}건",    C["green"]),
                    (" 실패", f"{fail:,}건",  C["red"]),
                    ("⏭ 패스",  f"{skip:,}건", C["muted"]),
                ]
                self.after(0,lambda k=kpis:self._geo_kpi_cards(k))
                time.sleep(delay)

            # ── 결과를 df_raw / df_view 에 반영 ──────────────────
            for c in RESULT_COLS:
                if c in df.columns:
                    self.df_raw[c]=df[c].values if len(df)==len(self.df_raw)                                    else df[c]
            if self.df_view is not None:
                for c in RESULT_COLS:
                    if c in df.columns: self.df_view[c]=df[c]
            # 컬럼 목록 갱신
            self.all_cols=list(self.df_raw.columns)
            self.vis_cols=list(self.all_cols)

            # 테이블 갱신
            self.after(0,lambda:self._render_tbl(self.df_raw))
            self.after(0,lambda:self._render_cols())

            done_msg=(f"\n 완료!  총: {n:,}건 | 성공: {ok:,}건 | "
                      f"실패: {fail:,}건 | 패스: {skip:,}건")
            self.after(0,lambda:self._geo_log(done_msg))
            self.after(0,lambda:self.geo_pb_var.set(100))
            self.after(0,lambda:self.geo_lbl_prog.config(text="완료"))
            self.after(0,lambda:messagebox.showinfo("완료",
                f"위경도 변환 완료!\n\n성공: {ok:,}건\n실패: {fail:,}건\n패스: {skip:,}건"))

        except Exception as e:
            msg=f" 오류: {str(e)}"
            self.after(0,lambda:self._geo_log(msg))
            self.after(0,lambda:messagebox.showerror("오류",str(e)))
        finally:
            self.after(0,lambda:self.geo_btn.config(state="normal"))


    # ======================== UNDO / REDO ========================
    def _push_undo(self, label="작업"):
        """현재 df 상태를 Undo 스택에 저장 (최대 20단계)"""
        if self.df_raw is None: return
        import copy
        self._undo_stack.append((label,
                                  self.df_raw.copy(),
                                  self.df_view.copy() if self.df_view is not None else None))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def _undo(self):
        if not self._undo_stack:
            self._st("되돌릴 작업이 없습니다", C["muted"]); return
        label, df_raw, df_view = self._undo_stack.pop()
        # 현재 상태를 redo에 저장
        if self.df_raw is not None:
            self._redo_stack.append(("redo",
                                     self.df_raw.copy(),
                                     self.df_view.copy() if self.df_view is not None else None))
        self.df_raw  = df_raw
        self.df_view = df_view
        self.all_cols = list(df_raw.columns)
        self.vis_cols = list(df_raw.columns)
        self._render_cols(); self._init_conds()
        self._render_tbl(df_view if df_view is not None else df_raw)
        self._upd_info()
        self._st(f"↺ 실행 취소: {label}", C["cyan"])

    def _redo(self):
        if not self._redo_stack:
            self._st("다시 실행할 작업이 없습니다", C["muted"]); return
        label, df_raw, df_view = self._redo_stack.pop()
        self._push_undo("redo")
        self.df_raw  = df_raw
        self.df_view = df_view
        self.all_cols = list(df_raw.columns)
        self.vis_cols = list(df_raw.columns)
        self._render_cols(); self._init_conds()
        self._render_tbl(df_view if df_view is not None else df_raw)
        self._upd_info()
        self._st(f"↻ 다시 실행", C["cyan"])

    # ======================== FILTER PRESETS ========================
    def _load_presets(self):
        import json
        try:
            with open(self._preset_file, encoding="utf-8") as f:
                self._filter_presets = json.load(f)
        except Exception:
            self._filter_presets = {}

    def _save_presets(self):
        import json
        try:
            with open(self._preset_file,"w",encoding="utf-8") as f:
                json.dump(self._filter_presets, f, ensure_ascii=False, indent=2)
        except Exception: pass

    def _save_preset(self):
        name = tk.simpledialog.askstring("조건 저장", "프리셋 이름 입력:",
                                          parent=self)
        if not name: return
        conds = [(cv.get(), self._opk(co), vv.get())
                 for cv,co,vv,_ in self.cond_rows if cv.get()]
        if not conds:
            messagebox.showwarning("주의","저장할 조건이 없습니다."); return
        self._filter_presets[name] = conds
        self._save_presets()
        self._refresh_preset_cmb()
        self._st(f"조건 저장: {name}", C["green"])

    def _load_preset(self):
        name = self.cmb_preset.get()
        if not name or name not in self._filter_presets:
            messagebox.showwarning("주의","프리셋을 선택하세요."); return
        conds = self._filter_presets[name]
        for _,_,_,fr in self.cond_rows: fr.destroy()
        self.cond_rows = []
        for col, op_key, val in conds:
            self._add_cond()
            cv,co,vv,_ = self.cond_rows[-1]
            if col in self.all_cols: cv.set(col)
            # op_key → 레이블 변환
            for lbl,k in OPS:
                if k == op_key: co.set(lbl); break
            vv.set(val)
        self._st(f"프리셋 로드: {name}", C["cyan"])

    def _del_preset(self):
        name = self.cmb_preset.get()
        if not name or name not in self._filter_presets: return
        if messagebox.askyesno("삭제 확인",f"'{name}' 프리셋을 삭제하시겠습니까?"):
            del self._filter_presets[name]
            self._save_presets()
            self._refresh_preset_cmb()
            self._st(f"프리셋 삭제: {name}", C["muted"])

    def _refresh_preset_cmb(self):
        try:
            self.cmb_preset["values"] = list(self._filter_presets.keys())
            if self._filter_presets:
                self.cmb_preset.set(list(self._filter_presets.keys())[-1])
            else:
                self.cmb_preset.set("")
        except Exception: pass

    # ======================== COLUMN ORDER ========================
    def _col_move(self, direction):
        """선택된 컬럼을 위/아래로 이동"""
        # 체크된 컬럼 중 마지막으로 클릭된 것 찾기 (간단히 첫 번째 체크 컬럼)
        checked = [c for c,v in self.col_vars.items() if v.get()]
        if not checked: return
        target = checked[0]
        idx = self.all_cols.index(target) if target in self.all_cols else -1
        if idx < 0: return
        new_idx = max(0, min(len(self.all_cols)-1, idx + direction))
        if new_idx == idx: return
        cols = list(self.all_cols)
        cols.insert(new_idx, cols.pop(idx))
        self.all_cols = cols
        self.vis_cols = [c for c in cols if self.col_vars.get(c,tk.BooleanVar(value=True)).get()]
        # df_raw 컬럼 순서 변경
        if self.df_raw is not None:
            existing = [c for c in cols if c in self.df_raw.columns]
            self.df_raw = self.df_raw[existing]
            if self.df_view is not None:
                self.df_view = self.df_view[[c for c in existing if c in self.df_view.columns]]
        self._render_cols()
        self._render_tbl(self.df_view if self.df_view is not None else self.df_raw)

    # ======================== COLUMN STATS CONTEXT MENU ========================
    def _col_ctx_menu(self, event):
        """테이블 헤더/셀 우클릭 → 컬럼 통계"""
        region = self.tree.identify_region(event.x, event.y)
        if region not in ("heading","cell"): return
        col_id = self.tree.identify_column(event.x)
        idx = int(col_id[1:])-1
        vis = [c for c in self.vis_cols if c in (self.df_view or self.df_raw or pd.DataFrame()).columns]
        if idx >= len(vis): return
        col = vis[idx]
        df  = self.df_view if self.df_view is not None else self.df_raw
        if df is None: return

        menu = tk.Menu(self, tearoff=0, bg=C["card"], fg=C["text"],
                       activebackground=C["accent"], activeforeground=C["bg"],
                       font=(FN,9))

        s = df[col]
        menu.add_command(label=f"컬럼: {col}", state="disabled")
        menu.add_separator()
        menu.add_command(label=f"건수:    {s.notna().sum():,}", state="disabled")
        menu.add_command(label=f"결측:    {s.isna().sum():,}", state="disabled")
        menu.add_command(label=f"유니크:  {s.nunique():,}", state="disabled")

        try:
            num = pd.to_numeric(s, errors="coerce").dropna()
            if len(num) > 0:
                menu.add_separator()
                menu.add_command(label=f"합계:    {num.sum():,.2f}", state="disabled")
                menu.add_command(label=f"평균:    {num.mean():,.2f}", state="disabled")
                menu.add_command(label=f"최솟값:  {num.min():,.2f}", state="disabled")
                menu.add_command(label=f"최댓값:  {num.max():,.2f}", state="disabled")
                menu.add_command(label=f"표준편차: {num.std():,.2f}", state="disabled")
        except Exception: pass

        menu.add_separator()
        menu.add_command(label="이 컬럼 오름차순 정렬",
                         command=lambda c=col:self._sort_by(c,True))
        menu.add_command(label="이 컬럼 내림차순 정렬",
                         command=lambda c=col:self._sort_by(c,False))
        menu.add_command(label="이 컬럼만 필터 조건으로 추가",
                         command=lambda c=col:self._add_cond_col(c))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _sort_by(self, col, asc):
        df = (self.df_view if self.df_view is not None else self.df_raw).copy()
        try:
            df[col] = pd.to_numeric(df[col], errors="ignore")
        except Exception: pass
        df = df.sort_values(col, ascending=asc, na_position="last")
        if self.df_view is not None: self.df_view = df
        else: self.df_raw = df
        self._render_tbl(df)
        self._st(f"{'오름' if asc else '내림'}차순: {col}", C["muted"])

    def _add_cond_col(self, col):
        self._add_cond()
        cv,co,vv,_ = self.cond_rows[-1]
        cv.set(col); co.set("포함")
        self.nb.select(0)

    # ======================== WRITE TO OPEN EXCEL ========================
    def _t_write(self):
        tab=ttk.Frame(self.nb)
        self.nb.add(tab,text="   엑셀 쓰기  ")
        fn=FN

        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="열린 Excel에 결과 직접 쓰기",
                 bg=C["card"],fg=C["text"],
                 font=(fn,11,"bold")).pack(anchor="w",padx=12,pady=(8,4))

        # 대상 파일/시트 선택
        r1=tk.Frame(cc,bg=C["card"]); r1.pack(fill="x",padx=12,pady=2)
        tk.Label(r1,text="대상 파일:",bg=C["card"],fg=C["muted"],font=(fn,9)).pack(side="left")
        self.cmb_wb=ttk.Combobox(r1,state="readonly",width=30)
        self.cmb_wb.pack(side="left",padx=4)
        self.cmb_wb.bind("<<ComboboxSelected>>",self._on_wb_pick)
        tk.Label(r1,text="시트:",bg=C["card"],fg=C["muted"],font=(fn,9)).pack(side="left",padx=(8,0))
        self.cmb_wsh=ttk.Combobox(r1,state="readonly",width=18)
        self.cmb_wsh.pack(side="left",padx=4)
        ttk.Button(r1,text=" 새로고침",style="Sm.TButton",
                   command=self._refresh_wb_list).pack(side="left",padx=4)

        # 쓰기 옵션
        r2=tk.Frame(cc,bg=C["card"]); r2.pack(fill="x",padx=12,pady=4)
        self.var_write_mode=tk.StringVar(value="overwrite")
        for val,txt in [("overwrite","기존 데이터 덮어쓰기"),
                        ("new_sheet","새 시트로 저장"),
                        ("append","기존 시트 아래에 추가")]:
            tk.Radiobutton(r2,text=txt,variable=self.var_write_mode,value=val,
                bg=C["card"],fg=C["text"],selectcolor=C["accent"],
                activebackground=C["card"],font=(fn,9)).pack(side="left",padx=8)

        r3=tk.Frame(cc,bg=C["card"]); r3.pack(fill="x",padx=12,pady=2)
        self.var_write_hdr=tk.BooleanVar(value=True)
        tk.Checkbutton(r3,text="헤더 포함",variable=self.var_write_hdr,
            bg=C["card"],fg=C["text"],selectcolor=C["accent"],
            activebackground=C["card"],font=(fn,9)).pack(side="left")
        self.var_write_filtered=tk.BooleanVar(value=True)
        tk.Checkbutton(r3,text="필터 결과만 (체크 해제 시 전체)",
            variable=self.var_write_filtered,
            bg=C["card"],fg=C["text"],selectcolor=C["accent"],
            activebackground=C["card"],font=(fn,9)).pack(side="left",padx=16)
        tk.Label(r3,text="시작 셀:",bg=C["card"],fg=C["muted"],font=(fn,9)).pack(side="left",padx=(16,4))
        self.var_start_cell=tk.StringVar(value="A1")
        ttk.Entry(r3,textvariable=self.var_start_cell,width=6).pack(side="left")

        bf=tk.Frame(cc,bg=C["card"]); bf.pack(fill="x",padx=12,pady=(6,10))
        ttk.Button(bf,text=" 엑셀에 쓰기",style="Teal.TButton",
                   command=self._do_write_excel).pack(side="left",ipadx=8)
        self.lbl_wr=tk.Label(bf,text="",bg=C["card"],fg=C["cyan"],font=(fn,9))
        self.lbl_wr.pack(side="left",padx=10)

        # 로그
        lf=tk.Frame(tab,bg=C["card"]); lf.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(lf,text="쓰기 로그",bg=C["card"],fg=C["text"],
                 font=(fn,10,"bold")).pack(anchor="w",padx=10,pady=(8,4))
        from tkinter import scrolledtext as _sc
        self.write_log=_sc.ScrolledText(lf,height=10,bg=C["card2"],fg=C["text2"],
                                        font=("Consolas" if IS_WIN else "Courier",9),
                                        relief="flat")
        self.write_log.pack(fill="both",expand=True,padx=8,pady=(0,8))

    def _refresh_wb_list(self):
        try:
            import xlwings as xw
            names=[]
            for _a in xw.apps:
                for _w in _a.books:
                    names.append(_w.name)
            self.cmb_wb["values"]=names
            if names: self.cmb_wb.set(names[0]); self._on_wb_pick()
            self._st(f"Excel 목록 새로고침: {len(names)}개",C["green"])
        except Exception as e:
            self._st(f"새로고침 실패: {e}",C["red"])

    def _on_wb_pick(self, event=None):
        try:
            import xlwings as xw
            nm=self.cmb_wb.get()
            for _a in xw.apps:
                for _w in _a.books:
                    if _w.name==nm:
                        sheets=[s.name for s in _w.sheets]
                        self.cmb_wsh["values"]=sheets
                        if sheets: self.cmb_wsh.set(sheets[0])
                        return
        except Exception: pass

    def _wlog(self,msg):
        self.write_log.insert(tk.END,msg+"\n")
        self.write_log.see(tk.END); self.update_idletasks()

    def _do_write_excel(self):
        if self.df_raw is None:
            messagebox.showwarning("주의","데이터를 먼저 불러오세요."); return
        wb_name=self.cmb_wb.get(); sh_name=self.cmb_wsh.get()
        if not wb_name:
            messagebox.showwarning("주의","대상 파일을 선택하세요.\n[ 새로고침] 클릭 후 선택"); return

        src=self.df_view if (self.var_write_filtered.get() and self.df_view is not None)             else self.df_raw
        out=src[[c for c in self.vis_cols if c in src.columns]]
        mode=self.var_write_mode.get()
        cell=self.var_start_cell.get().strip() or "A1"
        hdr=self.var_write_hdr.get()

        try:
            import xlwings as xw, re as _re
            self._wlog(f"대상: {wb_name} / {sh_name}  |  {len(out):,}행 × {len(out.columns)}열")
            # 파일 찾기
            xw_wb=None
            for _a in xw.apps:
                for _w in _a.books:
                    if _w.name==wb_name: xw_wb=_w; break
                if xw_wb: break
            if xw_wb is None:
                raise ValueError(f"'{wb_name}' 파일을 찾을 수 없습니다.")

            if mode=="new_sheet":
                sh_name=sh_name+"_결과" if sh_name else "결과"
                try: xw_wb.sheets.add(sh_name)
                except Exception: sh_name=f"결과_{pd.Timestamp.now().strftime('%H%M%S')}"
                xw_wb.sheets.add(sh_name)
                self._wlog(f"새 시트 생성: {sh_name}")

            ws=xw_wb.sheets[sh_name]

            if mode=="overwrite":
                ws.clear_contents()
                self._wlog("기존 내용 삭제 완료")

            # 시작 셀 파싱
            m=_re.match(r"([A-Za-z]+)(\d+)",cell)
            if not m: col_s,row_s=1,1
            else:
                col_str=m.group(1).upper()
                col_s=sum((ord(c)-64)*26**i
                          for i,c in enumerate(reversed(col_str)))
                row_s=int(m.group(2))

            if mode=="append":
                last_row=ws.range((ws.cells.last_cell.row,col_s)).end("up").row
                row_s=last_row+2 if last_row>1 else 1
                self._wlog(f"추가 위치: {row_s}행부터")

            # 헤더 쓰기
            if hdr:
                ws.range((row_s,col_s)).value=[list(out.columns)]
                row_s+=1

            # 데이터 쓰기 (None 변환)
            data=out.where(pd.notna(out),None).values.tolist()
            ws.range((row_s,col_s)).value=data

            self._wlog(f" 쓰기 완료: {len(out):,}행 → {wb_name} / {sh_name}")
            self.lbl_wr.config(text=f" {len(out):,}행 쓰기 완료")
            self._st(f"엑셀 쓰기 완료: {wb_name}",C["green"])
            messagebox.showinfo("완료",f"{wb_name}\n{sh_name}에\n{len(out):,}행을 저장했습니다.")

        except Exception as e:
            self._wlog(f" 오류: {e}")
            self.lbl_wr.config(text=f"오류: {e}")
            messagebox.showerror("쓰기 오류",str(e))

    # ======================== MULTI FILE MERGE ========================
    def _t_multmerge(self):
        tab=ttk.Frame(self.nb)
        self.nb.add(tab,text="   다중파일 병합  ")
        fn=FN

        cc=tk.Frame(tab,bg=C["card"]); cc.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(cc,text="다중 파일 일괄 병합 (동일 구조 Excel/CSV)",
                 bg=C["card"],fg=C["text"],
                 font=(fn,11,"bold")).pack(anchor="w",padx=12,pady=(8,4))

        # 파일 목록
        lf=tk.Frame(cc,bg=C["card"]); lf.pack(fill="x",padx=12,pady=2)
        tk.Label(lf,text="병합할 파일 목록:",bg=C["card"],fg=C["muted"],
                 font=(fn,8,"bold")).pack(anchor="w")

        lb_frame=tk.Frame(cc,bg=C["card"]); lb_frame.pack(fill="x",padx=12,pady=2)
        lb_sb=ttk.Scrollbar(lb_frame,orient="vertical")
        self.lb_mfiles=tk.Listbox(lb_frame,height=5,
            bg=C["card2"],fg=C["text2"],selectmode="multiple",
            font=(fn,9),yscrollcommand=lb_sb.set,
            selectbackground=C["accent"])
        lb_sb.config(command=self.lb_mfiles.yview)
        lb_sb.pack(side="right",fill="y")
        self.lb_mfiles.pack(fill="x",expand=True)

        bf2=tk.Frame(cc,bg=C["card"]); bf2.pack(fill="x",padx=12,pady=4)
        ttk.Button(bf2,text=" 파일 추가",style="Sm.TButton",
                   command=self._multmerge_add).pack(side="left")
        ttk.Button(bf2,text=" 폴더 추가",style="Sm.TButton",
                   command=self._multmerge_folder).pack(side="left",padx=4)
        ttk.Button(bf2,text="선택 삭제",style="Sm.TButton",
                   command=lambda:self.lb_mfiles.delete(
                       *self.lb_mfiles.curselection()[::-1])
                   ).pack(side="left")
        ttk.Button(bf2,text="전체 삭제",style="Flat.TButton",
                   command=lambda:self.lb_mfiles.delete(0,"end")
                   ).pack(side="left",padx=4)

        # 옵션
        r4=tk.Frame(cc,bg=C["card"]); r4.pack(fill="x",padx=12,pady=2)
        tk.Label(r4,text="헤더 행:",bg=C["card"],fg=C["muted"],font=(fn,9)).pack(side="left")
        self.spn_mhdr=tk.Spinbox(r4,from_=0,to=10,width=3,
            bg=C["card2"],fg=C["text"],buttonbackground=C["border"],
            relief="flat",font=(fn,9))
        self.spn_mhdr.delete(0,"end"); self.spn_mhdr.insert(0,"0")
        self.spn_mhdr.pack(side="left",padx=4)
        self.var_add_src=tk.BooleanVar(value=True)
        tk.Checkbutton(r4,text="소스파일명 컬럼 추가",variable=self.var_add_src,
            bg=C["card"],fg=C["text"],selectcolor=C["accent"],
            activebackground=C["card"],font=(fn,9)).pack(side="left",padx=12)
        self.var_skip_empty=tk.BooleanVar(value=True)
        tk.Checkbutton(r4,text="빈 행 제거",variable=self.var_skip_empty,
            bg=C["card"],fg=C["text"],selectcolor=C["accent"],
            activebackground=C["card"],font=(fn,9)).pack(side="left")

        bf3=tk.Frame(cc,bg=C["card"]); bf3.pack(fill="x",padx=12,pady=(6,10))
        ttk.Button(bf3,text=" 병합 실행",style="Teal.TButton",
                   command=self._do_multmerge).pack(side="left",ipadx=8)
        ttk.Button(bf3,text=" 결과 저장",style="Green.TButton",
                   command=self._save_multmerge).pack(side="left",padx=6)
        ttk.Button(bf3,text="현재 탭에 적용",style="Flat.TButton",
                   command=self._apply_multmerge).pack(side="left")
        self.lbl_mmr=tk.Label(bf3,text="",bg=C["card"],fg=C["cyan"],font=(fn,9))
        self.lbl_mmr.pack(side="left",padx=10)

        # 결과 미리보기
        tc=tk.Frame(tab,bg=C["card"]); tc.pack(fill="both",expand=True,padx=6,pady=(0,6))
        self.lbl_mm_tbl=tk.Label(tc,text="병합 결과 미리보기",
            bg=C["card"],fg=C["text"],font=(fn,10,"bold"))
        self.lbl_mm_tbl.pack(anchor="w",padx=10,pady=(8,4))
        tw=tk.Frame(tc,bg=C["card"]); tw.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.mm_tree=VTree(tw,show="headings")
        mvsb=ttk.Scrollbar(tw,orient="vertical",command=self.mm_tree.yview)
        mhsb=ttk.Scrollbar(tw,orient="horizontal",command=self.mm_tree.xview)
        self.mm_tree.configure(yscrollcommand=mvsb.set,xscrollcommand=mhsb.set)
        mhsb.pack(side="bottom",fill="x"); mvsb.pack(side="right",fill="y")
        self.mm_tree.pack(fill="both",expand=True)
        self._mm_result=None

    def _multmerge_add(self):
        paths=filedialog.askopenfilenames(
            title="병합할 파일 선택",
            filetypes=[("Excel/CSV","*.xlsx *.xls *.xlsm *.csv"),("All","*.*")])
        for p in paths:
            if p not in self.lb_mfiles.get(0,"end"):
                self.lb_mfiles.insert("end",p)

    def _multmerge_folder(self):
        import glob
        folder=filedialog.askdirectory(title="폴더 선택")
        if not folder: return
        for p in sorted(glob.glob(os.path.join(folder,"*.xlsx"))+
                        glob.glob(os.path.join(folder,"*.xls"))+
                        glob.glob(os.path.join(folder,"*.csv"))):
            if p not in self.lb_mfiles.get(0,"end"):
                self.lb_mfiles.insert("end",p)
        self._st(f"폴더 추가 완료",C["green"])

    def _do_multmerge(self):
        paths=list(self.lb_mfiles.get(0,"end"))
        if not paths:
            messagebox.showwarning("주의","파일을 추가하세요."); return
        try:
            hdr=int(self.spn_mhdr.get())
        except Exception: hdr=0

        dfs=[]; errors=[]
        for p in paths:
            try:
                ext=os.path.splitext(p)[1].lower()
                if ext==".csv":
                    for enc in ["utf-8-sig","euc-kr","cp949","utf-8"]:
                        try:
                            df=pd.read_csv(p,encoding=enc,header=hdr,low_memory=False)
                            break
                        except Exception: pass
                else:
                    eng="openpyxl" if ext in(".xlsx",".xlsm") else "xlrd"
                    df=pd.read_excel(p,header=hdr,engine=eng)

                if self.var_skip_empty.get():
                    df=df.dropna(how="all")
                if self.var_add_src.get():
                    df.insert(0,"__소스파일__",os.path.basename(p))
                dfs.append(df)
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")

        if not dfs:
            messagebox.showerror("오류","읽을 수 있는 파일이 없습니다.\n"+
                                 "\n".join(errors)); return

        merged=pd.concat(dfs,ignore_index=True)
        self._mm_result=merged
        self.mm_tree.load(merged,list(merged.columns))
        total=sum(len(d) for d in dfs)
        self.lbl_mm_tbl.config(
            text=f"병합 결과: {len(paths)}개 파일  |  {total:,}행 → {len(merged):,}행 (빈행제거 후)")
        self.lbl_mmr.config(text=f" {len(merged):,}행 병합 완료")
        if errors:
            messagebox.showwarning("일부 오류",
                f"{len(errors)}개 파일 읽기 실패:\n"+"\n".join(errors))
        self._st(f"다중 병합 완료: {len(merged):,}행",C["green"])

    def _save_multmerge(self):
        if self._mm_result is None:
            messagebox.showinfo("알림","병합을 먼저 실행하세요."); return
        self._save_df(self._mm_result)

    def _apply_multmerge(self):
        if self._mm_result is None:
            messagebox.showinfo("알림","병합을 먼저 실행하세요."); return
        self._apply(self._mm_result,"다중파일 병합")
        self.nb.select(0)
        self._st("다중 병합 결과 → 현재 데이터 적용",C["green"])

    # ======================== MATCH (VLOOKUP형 다중 조건 매칭) ========================
    def _t_match(self):
        tab=ttk.Frame(self.nb); self.nb.add(tab,text="  매칭  ")

        # ── A 파일 ─────────────────────────────────────────
        ac=tk.Frame(tab,bg=C["card"]); ac.pack(fill="x",padx=6,pady=(6,3))
        tk.Label(ac,text="A 파일  (기준 데이터)",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")).pack(anchor="w",padx=12,pady=(8,4))
        af=tk.Frame(ac,bg=C["card"]); af.pack(fill="x",padx=12,pady=(0,8))
        ttk.Button(af,text="현재 파일 사용",style="Flat.TButton",
                   command=self._match_use_current).pack(side="left")
        ttk.Button(af,text=" 찾아보기",style="Sm.TButton",
                   command=lambda:self._match_browse("a")).pack(side="left",padx=6)
        ttk.Button(af,text=" 열린 파일",style="Green.TButton",
                   command=lambda:self._match_pick_open("a")).pack(side="left")
        self.lbl_ma=tk.Label(af,text="(없음)",bg=C["card"],fg=C["cyan"],font=(FN,9))
        self.lbl_ma.pack(side="left",padx=8)

        # ── B 파일 ─────────────────────────────────────────
        bc=tk.Frame(tab,bg=C["card"]); bc.pack(fill="x",padx=6,pady=(0,3))
        tk.Label(bc,text="B 파일  (참조 / 조회)",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")).pack(anchor="w",padx=12,pady=(8,4))
        bf2=tk.Frame(bc,bg=C["card"]); bf2.pack(fill="x",padx=12,pady=(0,8))
        ttk.Button(bf2,text=" 찾아보기",style="Sm.TButton",
                   command=lambda:self._match_browse("b")).pack(side="left")
        ttk.Button(bf2,text=" 열린 파일",style="Green.TButton",
                   command=lambda:self._match_pick_open("b")).pack(side="left",padx=6)
        self.lbl_mb=tk.Label(bf2,text="(없음)",bg=C["card"],fg=C["cyan"],font=(FN,9))
        self.lbl_mb.pack(side="left",padx=8)
        tk.Label(bf2,text="시트:",bg=C["card"],fg=C["muted"],font=(FN,9)).pack(side="left",padx=(16,4))
        self.cmb_bsh=ttk.Combobox(bf2,state="readonly",width=16)
        self.cmb_bsh.pack(side="left")
        self.cmb_bsh.bind("<<ComboboxSelected>>",lambda e:self._match_b_sheet_chg())

        # ── 매칭 키 ────────────────────────────────────────
        mc=tk.Frame(tab,bg=C["card"]); mc.pack(fill="x",padx=6,pady=(0,3))
        mh=tk.Frame(mc,bg=C["card"]); mh.pack(fill="x",padx=12,pady=(8,4))
        tk.Label(mh,text="매칭 키  (다중 조건 가능)",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")).pack(side="left")
        ttk.Button(mh,text="＋ 키 추가",style="Flat.TButton",
                   command=self._match_add_key).pack(side="left",padx=12)
        kh=tk.Frame(mc,bg=C["card"]); kh.pack(fill="x",padx=12)
        tk.Label(kh,text="A 컬럼",bg=C["card"],fg=C["muted"],font=(FN,9),width=22,anchor="w").pack(side="left")
        tk.Label(kh,text="↔",bg=C["card"],fg=C["muted"],font=(FN,9),width=4).pack(side="left")
        tk.Label(kh,text="B 컬럼",bg=C["card"],fg=C["muted"],font=(FN,9),width=22,anchor="w").pack(side="left")
        self._match_key_frame=tk.Frame(mc,bg=C["card"])
        self._match_key_frame.pack(fill="x",padx=12,pady=(2,8))
        self._match_key_rows=[]  # [(a_var,b_var,a_cmb,b_cmb,frame), ...]

        # ── 가져올 컬럼 ────────────────────────────────────
        cc2=tk.Frame(tab,bg=C["card"]); cc2.pack(fill="x",padx=6,pady=(0,3))
        ch=tk.Frame(cc2,bg=C["card"]); ch.pack(fill="x",padx=12,pady=(8,4))
        tk.Label(ch,text="가져올 컬럼 선택  (B → A, 다중)",bg=C["card"],fg=C["text"],
                 font=(FN,11,"bold")).pack(side="left")
        ttk.Button(ch,text="전체 선택",style="Flat.TButton",
                   command=lambda:self.lb_match_cols.select_set(0,"end")).pack(side="left",padx=8)
        ttk.Button(ch,text="전체 해제",style="Flat.TButton",
                   command=lambda:self.lb_match_cols.selection_clear(0,"end")).pack(side="left")
        lbf2=tk.Frame(cc2,bg=C["card"]); lbf2.pack(fill="x",padx=12,pady=(0,8))
        sb4=ttk.Scrollbar(lbf2,orient="vertical")
        self.lb_match_cols=tk.Listbox(lbf2,selectmode="multiple",
                                      bg=C["card2"],fg=C["text2"],
                                      selectbackground=C["accent"],height=4,
                                      font=(FN,9),yscrollcommand=sb4.set)
        sb4.config(command=self.lb_match_cols.yview)
        sb4.pack(side="right",fill="y")
        self.lb_match_cols.pack(fill="x",expand=True)

        # ── 실행 ───────────────────────────────────────────
        xf=tk.Frame(tab,bg=C["card"]); xf.pack(fill="x",padx=6,pady=(0,3))
        xb=tk.Frame(xf,bg=C["card"]); xb.pack(fill="x",padx=12,pady=(8,10))
        ttk.Button(xb,text=" 매칭 실행",style="Teal.TButton",
                   command=self._do_match).pack(side="left")
        ttk.Button(xb,text="결과 저장",style="Green.TButton",
                   command=self._save_match).pack(side="left",padx=6)
        ttk.Button(xb,text="현재 데이터에 적용",style="Purple.TButton",
                   command=self._apply_match).pack(side="left")
        self.lbl_match_r=tk.Label(xb,text="",bg=C["card"],fg=C["cyan"],font=(FN,9))
        self.lbl_match_r.pack(side="left",padx=10)

        # ── 결과 미리보기 ──────────────────────────────────
        tc2=tk.Frame(tab,bg=C["card"]); tc2.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tk.Label(tc2,text="매칭 결과 미리보기",bg=C["card"],fg=C["text"],
                 font=(FN,10,"bold")).pack(anchor="w",padx=10,pady=(8,4))
        mw2=tk.Frame(tc2,bg=C["card"]); mw2.pack(fill="both",expand=True,padx=8,pady=(0,8))
        self.match_tree=VTree(mw2,show="headings")
        mvsb2=ttk.Scrollbar(mw2,orient="vertical",command=self.match_tree.yview)
        mhsb2=ttk.Scrollbar(mw2,orient="horizontal",command=self.match_tree.xview)
        self.match_tree.configure(yscrollcommand=mvsb2.set,xscrollcommand=mhsb2.set)
        mhsb2.pack(side="bottom",fill="x"); mvsb2.pack(side="right",fill="y")
        self.match_tree.pack(fill="both",expand=True)

        self._match_df_a=None; self._match_df_b=None
        self._match_result=None; self._match_b_xl=None

    def _match_pick_open(self, which):
        """열려있는 Excel 파일 목록 팝업 → A 또는 B로 로드"""
        books = self._get_open_books()
        if not books:
            messagebox.showwarning("열린 파일 없음",
                "Excel 파일이 열려 있지 않습니다.\n"
                "Excel에서 파일을 먼저 열어주세요.")
            return

        pop = tk.Toplevel(self)
        pop.title("열린 Excel 파일 선택")
        pop.configure(bg=C["surface"])
        pop.geometry("560x360")
        pop.resizable(False, False)
        pop.transient(self); pop.grab_set(); pop.lift(); pop.focus_force()

        lbl_title = "A 파일 (기준)" if which == "a" else "B 파일 (참조)"
        tk.Label(pop, text=f"{lbl_title}  —  열려있는 파일 선택",
                 bg=C["surface"], fg=C["text"],
                 font=(FN,11,"bold")).pack(pady=(14,6), padx=16, anchor="w")

        # 선택 상태를 변수로 직접 추적 (curselection 의존 제거)
        cur = {"bk": books[0] if books else None}

        # ── 파일 목록 (라디오버튼 스타일 프레임) ──────────────
        lf = tk.Frame(pop, bg=C["card"]); lf.pack(fill="x", padx=16)
        btn_vars = []
        lbl_sel_file = tk.Label(pop, text="", bg=C["surface"], fg=C["cyan"],
                                font=(FN,9)); lbl_sel_file.pack(anchor="w", padx=16)

        # 시트 콤보 먼저 생성 (파일 선택 버튼에서 참조)
        sf = tk.Frame(pop, bg=C["surface"]); sf.pack(fill="x", padx=16, pady=(8,4))
        tk.Label(sf, text="시트:", bg=C["surface"], fg=C["muted"],
                 font=(FN,9)).pack(side="left")
        cmb_sh2 = ttk.Combobox(sf, state="readonly", width=24)
        cmb_sh2.pack(side="left", padx=6)

        def _pick(bk):
            cur["bk"] = bk
            lbl_sel_file.config(text=f"  {bk['name']}  |  {bk['path']}")
            cmb_sh2["values"] = bk["sheets"]
            cmb_sh2.set(bk["sheets"][0] if bk["sheets"] else "")
            # 버튼 강조
            for v, btn in btn_vars:
                btn.config(relief="sunken" if v is bk else "flat",
                           fg=C["text"] if v is bk else C["muted"])

        for bk in books:
            row = tk.Frame(lf, bg=C["card2"], cursor="hand2")
            row.pack(fill="x", pady=1, padx=2)
            lbl = tk.Label(row,
                           text=f"  {bk['name']}    [{bk['source']}]    {bk['path']}",
                           bg=C["card2"], fg=C["muted"], font=(FN,9),
                           anchor="w", padx=4, pady=6)
            lbl.pack(fill="x")
            lbl.bind("<Button-1>", lambda e, b=bk: _pick(b))
            row.bind("<Button-1>",  lambda e, b=bk: _pick(b))
            btn_vars.append((bk, lbl))

        # 첫 번째 파일 기본 선택
        if books:
            _pick(books[0])

        def _load():
            bk = cur["bk"]
            if bk is None:
                messagebox.showwarning("주의","파일을 선택하세요.",parent=pop); return
            sh = cmb_sh2.get() or (bk["sheets"][0] if bk["sheets"] else "Sheet1")
            try:
                fp = bk.get("path","")
                df = None
                if fp and os.path.exists(fp):
                    ext = os.path.splitext(fp)[1].lower()
                    if ext in (".xlsx",".xlsm"):
                        df = pd.ExcelFile(fp, engine="openpyxl").parse(sh)
                    elif ext == ".xls":
                        df = pd.ExcelFile(fp, engine="xlrd").parse(sh)
                if df is None and bk.get("wb"):
                    ws = bk["wb"].sheets[sh]
                    data = ws.used_range.value
                    if data and isinstance(data[0], list):
                        df = pd.DataFrame(data[1:], columns=data[0])
                    elif data:
                        df = pd.DataFrame(data)
                if df is None:
                    raise ValueError("파일을 읽을 수 없습니다.\n경로: " + fp)
                df = self._clean(df)
                if which == "a":
                    self._match_df_a = df
                    self.lbl_ma.config(text=f"{bk['name']} / {sh}  ({len(df):,}행)")
                    self._match_refresh_combos("a")
                else:
                    self._match_df_b = df
                    self.lbl_mb.config(text=f"{bk['name']} / {sh}  ({len(df):,}행)")
                    self.cmb_bsh["values"] = bk["sheets"]
                    self.cmb_bsh.set(sh)
                    self._match_refresh_combos("b")
                    self.lb_match_cols.delete(0,"end")
                    for c in df.columns: self.lb_match_cols.insert("end",c)
                pop.destroy()
                self._st(f"매칭 {'A' if which=='a' else 'B'} 파일 로드: {bk['name']} / {sh}", C["teal"])
            except Exception as e:
                messagebox.showerror("오류",str(e),parent=pop)

        bf = tk.Frame(pop, bg=C["surface"]); bf.pack(pady=10)
        ttk.Button(bf, text="  선택  ", command=_load).pack(side="left", padx=6, ipadx=8, ipady=4)
        ttk.Button(bf, text="취소", style="Flat.TButton",
                   command=pop.destroy).pack(side="left", padx=6)
        ttk.Button(bf, text=" 새로고침", style="Flat.TButton",
                   command=lambda: (pop.destroy(),
                                    self.after(80, lambda: self._match_pick_open(which)))
                   ).pack(side="left", padx=6)

    def _match_use_current(self):
        if self.df_raw is None:
            messagebox.showwarning("주의","먼저 파일을 불러오세요."); return
        self._match_df_a=self.df_view if self.df_view is not None else self.df_raw
        name=os.path.basename(self.raw_path) if self.raw_path else "현재 파일"
        self.lbl_ma.config(text=f"{name}  ({len(self._match_df_a):,}행)")
        self._match_refresh_combos("a")

    def _match_browse(self,which):
        p=filedialog.askopenfilename(
            filetypes=[("Excel/CSV","*.xlsx *.xls *.csv"),("All","*.*")])
        if not p: return
        try:
            ext=os.path.splitext(p)[1].lower()
            if which=="a":
                if ext in(".xlsx",".xlsm"):   df=pd.ExcelFile(p,engine="openpyxl").parse(0)
                elif ext==".xls":              df=pd.ExcelFile(p,engine="xlrd").parse(0)
                else:                          df=Loader._csv(p,"utf-8",0)
                self._match_df_a=self._clean(df)
                self.lbl_ma.config(text=f"{os.path.basename(p)}  ({len(df):,}행)")
                self._match_refresh_combos("a")
            else:
                if ext in(".xlsx",".xlsm"):
                    xl=pd.ExcelFile(p,engine="openpyxl")
                    self._match_b_xl=xl
                    self.cmb_bsh["values"]=xl.sheet_names
                    self.cmb_bsh.set(xl.sheet_names[0])
                    df=xl.parse(xl.sheet_names[0])
                elif ext==".xls":
                    xl=pd.ExcelFile(p,engine="xlrd")
                    self._match_b_xl=xl
                    self.cmb_bsh["values"]=xl.sheet_names
                    self.cmb_bsh.set(xl.sheet_names[0])
                    df=xl.parse(xl.sheet_names[0])
                else:
                    self._match_b_xl=None
                    self.cmb_bsh["values"]=["Sheet1"]; self.cmb_bsh.set("Sheet1")
                    df=Loader._csv(p,"utf-8",0)
                self._match_df_b=self._clean(df)
                self.lbl_mb.config(text=f"{os.path.basename(p)}  ({len(df):,}행)")
                self._match_refresh_combos("b")
                self.lb_match_cols.delete(0,"end")
                for c in df.columns: self.lb_match_cols.insert("end",c)
        except Exception as e: messagebox.showerror("오류",str(e))

    def _match_b_sheet_chg(self):
        if self._match_b_xl is None: return
        try:
            df=self._match_b_xl.parse(self.cmb_bsh.get())
            self._match_df_b=self._clean(df)
            self.lbl_mb.config(text=f"{self.cmb_bsh.get()}  ({len(df):,}행)")
            self._match_refresh_combos("b")
            self.lb_match_cols.delete(0,"end")
            for c in df.columns: self.lb_match_cols.insert("end",c)
        except Exception as e: messagebox.showerror("오류",str(e))

    def _match_refresh_combos(self,which):
        a_cols=list(self._match_df_a.columns) if self._match_df_a is not None else []
        b_cols=list(self._match_df_b.columns) if self._match_df_b is not None else []
        for av,bv,ca,cb,_ in self._match_key_rows:
            if which=="a":
                ca["values"]=a_cols
                if av.get() not in a_cols and a_cols: av.set(a_cols[0])
            else:
                cb["values"]=b_cols
                if bv.get() not in b_cols and b_cols: bv.set(b_cols[0])

    def _match_add_key(self):
        fr=tk.Frame(self._match_key_frame,bg=C["card"]); fr.pack(fill="x",pady=2)
        a_cols=list(self._match_df_a.columns) if self._match_df_a is not None else []
        b_cols=list(self._match_df_b.columns) if self._match_df_b is not None else []
        av=tk.StringVar(); bv=tk.StringVar()
        ca=ttk.Combobox(fr,textvariable=av,state="readonly",width=22,values=a_cols)
        ca.pack(side="left")
        if a_cols: ca.set(a_cols[0])
        tk.Label(fr,text=" ↔ ",bg=C["card"],fg=C["muted"],font=(FN,9)).pack(side="left")
        cb=ttk.Combobox(fr,textvariable=bv,state="readonly",width=22,values=b_cols)
        cb.pack(side="left")
        if b_cols: cb.set(b_cols[0])
        row=(av,bv,ca,cb,fr)
        self._match_key_rows.append(row)
        def _del(r=row):
            self._match_key_rows.remove(r); r[4].destroy()
        ttk.Button(fr,text=" ✕ ",style="Flat.TButton",command=_del).pack(side="left",padx=4)

    def _do_match(self):
        if self._match_df_a is None:
            messagebox.showwarning("주의","A 파일을 불러오세요."); return
        if self._match_df_b is None:
            messagebox.showwarning("주의","B 파일을 불러오세요."); return
        if not self._match_key_rows:
            messagebox.showwarning("주의","매칭 키를 1개 이상 추가하세요."); return
        sel=self.lb_match_cols.curselection()
        if not sel:
            messagebox.showwarning("주의","가져올 컬럼을 선택하세요."); return
        keys_a=[av.get() for av,bv,_,_,_ in self._match_key_rows]
        keys_b=[bv.get() for av,bv,_,_,_ in self._match_key_rows]
        bring=[self.lb_match_cols.get(i) for i in sel]
        try:
            a=self._match_df_a.copy()
            need_b=list(dict.fromkeys(keys_b+bring))
            b=self._match_df_b[[c for c in need_b if c in self._match_df_b.columns]].copy()
            b=b.drop_duplicates(subset=keys_b,keep="first")
            b["__hit__"]=True
            merged=a.merge(b,left_on=keys_a,right_on=keys_b,how="left",suffixes=("","_B"))
            matched=int(merged["__hit__"].notna().sum())
            merged=merged.drop(columns=["__hit__"])
            # 중복 키 컬럼 제거 (B키가 A키와 이름이 다를 때)
            for ka,kb in zip(keys_a,keys_b):
                if ka!=kb and kb in merged.columns:
                    merged=merged.drop(columns=[kb])
            self._match_result=merged
            self.match_tree.load(merged,list(merged.columns))
            unmatched=len(merged)-matched
            self.lbl_match_r.config(
                text=f"총 {len(merged):,}행  |  매칭 {matched:,}건  |  미매칭 {unmatched:,}건")
            self._st(f"매칭 완료: {matched:,}/{len(merged):,}행",C["green"])
        except Exception as e: messagebox.showerror("매칭 오류",str(e))

    def _save_match(self):
        if self._match_result is None:
            messagebox.showinfo("알림","먼저 매칭을 실행하세요."); return
        self._save_df(self._match_result)

    def _apply_match(self):
        if self._match_result is None:
            messagebox.showinfo("알림","먼저 매칭을 실행하세요."); return
        self._apply(self._match_result,"매칭 결과 적용")
        self.nb.select(0)
        self._st("매칭 결과 → 현재 데이터 적용",C["green"])

    # ======================== CLOSE ========================
    def _close(self):
        plt.close("all"); self.destroy()

if __name__=="__main__":
    app=App(); app.mainloop()
