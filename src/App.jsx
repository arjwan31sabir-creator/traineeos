import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";

const CLAUDE_KEY  = import.meta.env.VITE_CLAUDE_API_KEY;
const WORK_LAT    = 23.5896598;
const WORK_LNG    = 58.4125277;
const WORK_RADIUS = 1000;
const MAX_SIGNIN  = "09:00";
const PENALTY_PCT = 8.33;

const KRA_CATEGORIES = [
  { id:"attendance",  label:"Attendance & Punctuality",  icon:"⏰", color:"#CF0A2C" },
  { id:"technical",   label:"Technical Skills",          icon:"🔧", color:"#FFA500" },
  { id:"reporting",   label:"Reporting & Documentation", icon:"📋", color:"#34d399" },
  { id:"teamwork",    label:"Teamwork & Communication",  icon:"🤝", color:"#4f8ef7" },
  { id:"learning",    label:"Learning & Development",    icon:"📚", color:"#7c5cfc" },
  { id:"initiative",  label:"Initiative & Innovation",   icon:"💡", color:"#FFD700" },
];

const CHINESE_QUOTES = [
  { chinese:"千里之行，始于足下", english:"A journey of a thousand miles begins with a single step.", author:"老子 Lao Tzu" },
  { chinese:"学而不思则罔，思而不学则殆", english:"Learning without thought is labor lost; thought without learning is perilous.", author:"孔子 Confucius" },
  { chinese:"不积跬步，无以至千里", english:"Without accumulating small steps, one cannot walk a thousand miles.", author:"荀子 Xunzi" },
  { chinese:"志不强者智不达", english:"Those with weak ambition cannot achieve great wisdom.", author:"墨子 Mozi" },
  { chinese:"勤能补拙", english:"Diligence can make up for lack of talent.", author:"中国谚语 Chinese Proverb" },
  { chinese:"宝剑锋从磨砺出，梅花香自苦寒来", english:"A sharp sword comes from grinding; the fragrance of plum blossoms comes from bitter cold.", author:"中国谚语 Chinese Proverb" },
  { chinese:"失败是成功之母", english:"Failure is the mother of success.", author:"中国谚语 Chinese Proverb" },
  { chinese:"活到老，学到老", english:"Live and learn until old age.", author:"中国谚语 Chinese Proverb" },
  { chinese:"知己知彼，百战不殆", english:"Know yourself and your enemy, and you will never be defeated.", author:"孙子 Sun Tzu" },
  { chinese:"人无远虑，必有近忧", english:"He who does not think ahead will find trouble at his doorstep.", author:"孔子 Confucius" },
];

const THEMES = [
  { name:"Huawei Red",    primary:"#CF0A2C", darkPrimary:"#A00820", dark:"#0D0D0D", surface:"#1E1E1E", surface2:"#2A2A2A", border:"#333333", text:"#F5F5F5", muted:"#888888", white:"#FFFFFF" },
  { name:"Ocean Blue",    primary:"#0066FF", darkPrimary:"#0044BB", dark:"#050E1F", surface:"#0A1628", surface2:"#0F2040", border:"#1A3A6A", text:"#F0F8FF", muted:"#7BA7D4", white:"#FFFFFF" },
  { name:"Royal Purple",  primary:"#7C3AED", darkPrimary:"#5B21B6", dark:"#0D0A1F", surface:"#150E2E", surface2:"#1E1440", border:"#2D1F60", text:"#F5F0FF", muted:"#A78BCA", white:"#FFFFFF" },
  { name:"Emerald Green", primary:"#059669", darkPrimary:"#047857", dark:"#030F0A", surface:"#071A10", surface2:"#0A2518", border:"#0F3D26", text:"#F0FFF8", muted:"#6BB891", white:"#FFFFFF" },
  { name:"Sunset Orange", primary:"#EA580C", darkPrimary:"#C2410C", dark:"#100704", surface:"#1C0D06", surface2:"#28140A", border:"#4A200E", text:"#FFF7F0", muted:"#C48A6A", white:"#FFFFFF" },
  { name:"Rose Pink",     primary:"#E11D78", darkPrimary:"#BE1065", dark:"#0F040A", surface:"#1C0712", surface2:"#28091A", border:"#4A1030", text:"#FFF0F7", muted:"#C47A9E", white:"#FFFFFF" },
  { name:"Sky Teal",      primary:"#0891B2", darkPrimary:"#0E7490", dark:"#030D10", surface:"#071820", surface2:"#0A2230", border:"#0F3A4A", text:"#F0FEFF", muted:"#60A8BB", white:"#FFFFFF" },
  { name:"Golden Sun",    primary:"#D97706", darkPrimary:"#B45309", dark:"#0F0A02", surface:"#1C1404", surface2:"#281E06", border:"#4A380A", text:"#FFFBF0", muted:"#BBA050", white:"#FFFFFF" },
];

const randomTheme = THEMES[Math.floor(Math.random()*THEMES.length)];
const HW = {
  red:randomTheme.primary, darkRed:randomTheme.darkPrimary,
  black:"#1A1A1A", dark:randomTheme.dark,
  surface:randomTheme.surface, surface2:randomTheme.surface2,
  border:randomTheme.border, text:randomTheme.text,
  muted:randomTheme.muted, white:randomTheme.white,
};

function getCurrentWeek() {
  const now = new Date();
  const day = now.getDay();
  const sunday = new Date(now);
  sunday.setDate(now.getDate() - day);
  const thursday = new Date(sunday);
  thursday.setDate(sunday.getDate() + 4);
  const fmt = (d) => d.toISOString().split("T")[0];
  return {
    week_start: fmt(sunday),
    week_end: fmt(thursday),
    label: `${sunday.toLocaleDateString("en-GB",{day:"numeric",month:"short"})} – ${thursday.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})}`,
  };
}

function HuaweiLogo({size=32}) {
  return (
    <svg width={size} height={size} viewBox="0 0 100 100">
      {[0,60,120,180,240,300].map(r=>(
        <ellipse key={r} cx="50" cy="22" rx="8" ry="20" fill={HW.red}
          transform={`rotate(${r} 50 50)`}/>
      ))}
    </svg>
  );
}

function getDistance(lat1,lng1,lat2,lng2) {
  const R=6371000,dLat=(lat2-lat1)*Math.PI/180,dLng=(lng2-lng1)*Math.PI/180;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*
    Math.cos(lat2*Math.PI/180)*Math.sin(dLng/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}

async function analyzeReport(text) {
  const res=await fetch("https://api.anthropic.com/v1/messages",{
    method:"POST",
    headers:{"Content-Type":"application/json","x-api-key":CLAUDE_KEY,
      "anthropic-version":"2023-06-01",
      "anthropic-dangerous-direct-browser-access":"true"},
    body:JSON.stringify({model:"claude-haiku-4-5-20251001",max_tokens:1024,
      messages:[{role:"user",content:
        `Analyze this trainee weekly report and respond ONLY with JSON:
Report: "${text}"
{"kpi_score":<0-100>,"pie_chart":{"Tasks Completed":<pct>,"Planning":<pct>,"Challenges":<pct>,"Learning":<pct>},"talent_notes":"<2-3 sentences>","summary":"<one sentence>"}`
      }]})
  });
  const data=await res.json();
  return JSON.parse(data.content[0].text.replace(/```json|```/g,"").trim());
}

function PieChart({data}) {
  const colors=[HW.red,"#FF6B6B","#FF9999","#FFB3B3"];
  const entries=Object.entries(data);
  let cum=0;
  const slices=entries.map(([label,pct],i)=>{
    const val=pct/100,s=cum*2*Math.PI;cum+=val;const e=cum*2*Math.PI;
    const x1=Math.cos(s-Math.PI/2),y1=Math.sin(s-Math.PI/2);
    const x2=Math.cos(e-Math.PI/2),y2=Math.sin(e-Math.PI/2);
    return{label,pct,color:colors[i],
      d:`M 0 0 L ${x1} ${y1} A 1 1 0 ${val>.5?1:0} 1 ${x2} ${y2} Z`};
  });
  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:12}}>
      <svg viewBox="-1.1 -1.1 2.2 2.2" width="140" height="140">
        {slices.map((s,i)=>(
          <path key={i} d={s.d} fill={s.color} stroke={HW.surface} strokeWidth="0.03"/>
        ))}
        <circle cx="0" cy="0" r="0.55" fill={HW.surface}/>
      </svg>
      <div style={{display:"flex",flexWrap:"wrap",gap:8,justifyContent:"center"}}>
        {slices.map((s,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:6,fontSize:12}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:s.color,flexShrink:0}}/>
            <span style={{color:HW.muted}}>{s.label}</span>
            <span style={{fontWeight:700,color:s.color}}>{s.pct}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function OKRBar({okr,onUpdate}) {
  const pct=Math.min((okr.current/okr.target)*100,100).toFixed(0);
  const color=pct>=80?HW.red:pct>=50?"#FFA500":"#666";
  return (
    <div style={{background:HW.surface2,borderRadius:12,padding:14,marginBottom:12}}>
      <div style={{marginBottom:6}}>
        <div style={{fontSize:11,color:HW.muted,marginBottom:2}}>{okr.department}</div>
        <div style={{fontWeight:600,fontSize:14,marginBottom:2}}>{okr.objective}</div>
        <div style={{fontSize:12,color:HW.muted}}>{okr.key_result}</div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}>
        <span style={{color:HW.muted}}>Progress</span>
        <span style={{fontWeight:700,color}}>{pct}% ({okr.current}/{okr.target} {okr.unit})</span>
      </div>
      <div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden",marginBottom:8}}>
        <div style={{height:"100%",borderRadius:10,background:color,
          width:`${pct}%`,transition:"width .6s ease"}}/>
      </div>
      {onUpdate&&(
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <input type="number" defaultValue={okr.current}
            style={{background:HW.surface,border:`1px solid ${HW.border}`,
              color:HW.text,borderRadius:6,padding:"6px 8px",flex:1,fontSize:14}}
            onBlur={e=>onUpdate(okr.id,parseFloat(e.target.value))}/>
          <span style={{fontSize:12,color:HW.muted}}>Update</span>
        </div>
      )}
    </div>
  );
}

function GoalCard({goal,onUpdate,onDelete,isTrainee}) {
  const kra=KRA_CATEGORIES.find(k=>k.id===goal.kra)||KRA_CATEGORIES[0];
  const pct=Math.min((goal.current_value/goal.target_value)*100,100).toFixed(0);
  const isOverdue=goal.due_date&&new Date(goal.due_date)<new Date()&&goal.status!=="completed";
  const status=isOverdue&&goal.status!=="completed"?"overdue":goal.status;
  const statusColors={
    not_started:{bg:"rgba(136,136,136,.15)",color:"#888"},
    in_progress:{bg:"rgba(79,142,247,.15)",color:"#4f8ef7"},
    completed:{bg:"rgba(52,211,153,.15)",color:"#34d399"},
    overdue:{bg:"rgba(248,113,113,.15)",color:"#f87171"},
  };
  return (
    <div style={{background:HW.surface2,borderRadius:14,padding:16,
      marginBottom:12,borderLeft:`4px solid ${kra.color}`}}>
      <div style={{display:"flex",justifyContent:"space-between",
        alignItems:"flex-start",marginBottom:10}}>
        <div style={{flex:1,marginRight:8}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
            <span style={{fontSize:16}}>{kra.icon}</span>
            <span style={{fontSize:10,color:kra.color,fontWeight:700,
              textTransform:"uppercase"}}>{kra.label}</span>
          </div>
          <div style={{fontWeight:700,fontSize:14,marginBottom:2}}>{goal.goal_title}</div>
          {goal.description&&<div style={{fontSize:12,color:HW.muted}}>{goal.description}</div>}
        </div>
        <span style={{padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,
          whiteSpace:"nowrap",background:statusColors[status]?.bg,
          color:statusColors[status]?.color}}>
          {status==="not_started"?"⬜ Not Started":status==="in_progress"?"🔵 In Progress":
           status==="completed"?"✅ Done":"🔴 Overdue"}
        </span>
      </div>
      <div style={{marginBottom:10}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
          <span style={{color:HW.muted}}>Progress</span>
          <span style={{fontWeight:700,color:kra.color}}>
            {goal.current_value}/{goal.target_value} {goal.unit} ({pct}%)
          </span>
        </div>
        <div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden"}}>
          <div style={{height:"100%",borderRadius:10,background:kra.color,
            width:`${pct}%`,transition:"width .6s ease"}}/>
        </div>
      </div>
      {isTrainee&&goal.status!=="completed"&&(
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <input type="number" placeholder="New value" id={`prog_${goal.id}`}
            style={{background:HW.surface,border:`1px solid ${HW.border}`,
              color:HW.text,borderRadius:6,padding:"8px 10px",flex:1,fontSize:14}}/>
          <button style={{padding:"8px 14px",borderRadius:6,border:"none",
            background:kra.color,color:HW.white,fontWeight:700,fontSize:13,cursor:"pointer"}}
            onClick={()=>{
              const input=document.getElementById(`prog_${goal.id}`);
              if(input&&input.value) onUpdate(goal.id,parseFloat(input.value));
            }}>Update</button>
          {parseFloat(pct)>=100&&(
            <button style={{padding:"8px 14px",borderRadius:6,border:"none",
              background:"rgba(52,211,153,.15)",color:"#34d399",
              fontWeight:700,fontSize:13,cursor:"pointer"}}
              onClick={()=>onUpdate(goal.id,goal.target_value,"completed")}>
              ✅ Complete
            </button>
          )}
          {onDelete&&(
            <button style={{padding:"8px 14px",borderRadius:6,border:"none",
              background:"rgba(248,113,113,.1)",color:"#f87171",
              fontWeight:700,fontSize:13,cursor:"pointer"}}
              onClick={()=>onDelete(goal.id)}>🗑</button>
          )}
        </div>
      )}
    </div>
  );
}

function ReminderPopup({onDismiss}) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",
      display:"flex",alignItems:"flex-end",justifyContent:"center",
      zIndex:999,backdropFilter:"blur(4px)"}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,
        borderRadius:"24px 24px 0 0",padding:32,width:"100%",
        maxWidth:480,textAlign:"center"}}>
        <div style={{fontSize:48,marginBottom:12}}>⏰</div>
        <h3 style={{color:HW.red,fontSize:20,marginBottom:8}}>Weekly Report Reminder</h3>
        <p style={{color:HW.muted,fontSize:14,lineHeight:1.6,marginBottom:20}}>
          Don't forget to submit your weekly report before Thursday!
        </p>
        <button onClick={onDismiss} style={{background:HW.red,color:HW.white,border:"none",
          borderRadius:12,padding:"14px 32px",fontWeight:700,fontSize:16,
          cursor:"pointer",width:"100%"}}>Got it!</button>
      </div>
    </div>
  );
}

function GreetingPopup({name,onDismiss}) {
  const now=new Date();
  const hour=now.getHours();
  const greeting=hour<12?"Good morning":hour<17?"Good afternoon":"Good evening";
  const greetingAr=hour<12?"صباح الخير":hour<17?"مساء الخير":"مساء النور";
  const quote=CHINESE_QUOTES[Math.floor(Math.random()*CHINESE_QUOTES.length)];
  useEffect(()=>{
    if("speechSynthesis" in window){
      window.speechSynthesis.cancel();
      const msg=new SpeechSynthesisUtterance(
        `${greeting}, ${name}! Welcome to Huawei TechTrack. ${quote.english}. Have a nice day!`
      );
      msg.rate=0.9;msg.pitch=1;msg.volume=1;
      const setVoice=()=>{
        const voices=window.speechSynthesis.getVoices();
        const preferred=voices.find(v=>v.lang.startsWith("en")&&v.name.includes("Female"))
          ||voices.find(v=>v.lang.startsWith("en"))||voices[0];
        if(preferred) msg.voice=preferred;
        setTimeout(()=>window.speechSynthesis.speak(msg),500);
      };
      if(window.speechSynthesis.getVoices().length>0) setVoice();
      else window.speechSynthesis.onvoiceschanged=setVoice;
    }
    return ()=>window.speechSynthesis?.cancel();
  },[]);
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.9)",
      display:"flex",alignItems:"center",justifyContent:"center",
      zIndex:1000,backdropFilter:"blur(8px)",padding:16}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,
        borderRadius:24,padding:32,width:"100%",maxWidth:420,
        textAlign:"center",animation:"popIn .4s cubic-bezier(.34,1.56,.64,1)",
        maxHeight:"90vh",overflowY:"auto"}}>
        <style>{`
          @keyframes popIn{from{transform:scale(.8);opacity:0}to{transform:scale(1);opacity:1}}
          @keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
        `}</style>
        <div style={{animation:"float 3s ease-in-out infinite",marginBottom:12}}>
          <HuaweiLogo size={56}/>
        </div>
        <div style={{fontSize:12,color:HW.muted,marginBottom:4,
          letterSpacing:".1em",textTransform:"uppercase"}}>{greetingAr}</div>
        <h2 style={{fontSize:28,fontWeight:800,color:HW.text,margin:"0 0 2px"}}>{greeting},</h2>
        <h2 style={{fontSize:32,fontWeight:800,color:HW.red,margin:"0 0 20px"}}>{name}! 👋</h2>
        <div style={{height:1,background:HW.border,marginBottom:20}}/>
        <div style={{background:HW.surface2,borderRadius:14,padding:16,
          marginBottom:16,border:`1px solid ${HW.border}`}}>
          <div style={{fontSize:10,color:HW.red,fontWeight:700,
            textTransform:"uppercase",letterSpacing:".1em",marginBottom:10}}>
            🈲 Today's Motivation
          </div>
          <div style={{fontSize:22,fontWeight:800,color:HW.text,marginBottom:8,
            lineHeight:1.4,fontFamily:"serif",letterSpacing:3}}>{quote.chinese}</div>
          <div style={{fontSize:13,color:HW.muted,lineHeight:1.6,
            fontStyle:"italic",marginBottom:6}}>"{quote.english}"</div>
          <div style={{fontSize:11,color:HW.red,fontWeight:600}}>— {quote.author}</div>
        </div>
        <div style={{fontSize:15,color:"#34d399",fontWeight:700,marginBottom:16}}>
          🌟 Have a nice day!
        </div>
        <div style={{fontSize:13,color:HW.muted,marginBottom:24}}>
          {now.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",
            month:"long",year:"numeric"})}
        </div>
        <div style={{fontSize:12,color:HW.red,fontWeight:700,marginBottom:16,
          background:`${HW.red}15`,borderRadius:8,padding:"6px 12px",
          display:"inline-block"}}>
          🎨 {randomTheme.name} Theme
        </div>
        <button onClick={onDismiss}
          style={{background:HW.red,color:HW.white,border:"none",
            borderRadius:12,padding:"16px 40px",fontWeight:800,
            fontSize:16,cursor:"pointer",width:"100%"}}>
          Let's Go! 🚀
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [view,setView]               = useState("login");
  const [traineeTab,setTraineeTab]   = useState("attendance");
  const [mgmtTab,setMgmtTab]         = useState("trainees");
  const [user,setUser]               = useState(null);
  const [traineeId,setTraineeId]     = useState(null);
  const [trainees,setTrainees]       = useState([]);
  const [selected,setSelected]       = useState(null);
  const [reports,setReports]         = useState([]);
  const [logs,setLogs]               = useState([]);
  const [penalties,setPenalties]     = useState([]);
  const [goals,setGoals]             = useState([]);
  const [selectedGoals,setSelectedGoals] = useState([]);
  const [okrs,setOkrs]               = useState([]);
  const [allReports,setAllReports]   = useState([]);
  const [allPenalties,setAllPenalties] = useState([]);
  const [liveSignins,setLiveSignins] = useState([]);
  const [loading,setLoading]         = useState(false);
  const [aiLoading,setAiLoading]     = useState(false);
  const [msg,setMsg]                 = useState("");
  const [email,setEmail]             = useState("");
  const [password,setPassword]       = useState("");
  const [confirmPwd,setConfirmPwd]   = useState("");
  const [profileTab,setProfileTab]   = useState("timeline");
  const [aiResult,setAiResult]       = useState(null);
  const [weeklyPhotoFile,setWeeklyPhotoFile] = useState(null);
  const [weeklyPhotoPreview,setWeeklyPhotoPreview] = useState(null);
  const [geoStatus,setGeoStatus]     = useState("idle");
  const [geoMsg,setGeoMsg]           = useState("");
  const [locationOk,setLocationOk]   = useState(false);
  const [showReminder,setShowReminder] = useState(false);
  const [showGreeting,setShowGreeting] = useState(false);
  const [traineeName,setTraineeName] = useState("");
  const [trafficCount,setTrafficCount] = useState(0);
  const [isLate,setIsLate]           = useState(false);
  const [excuseType,setExcuseType]   = useState("");
  const [excuseText,setExcuseText]   = useState("");
  const [excusePhoto,setExcusePhoto] = useState(null);
  const [excusePreview,setExcusePreview] = useState(null);
  const [newOkr,setNewOkr]           = useState({department:"",objective:"",
    key_result:"",target:100,current:0,unit:"%",due_date:""});
  const [showAddOkr,setShowAddOkr]   = useState(false);
  const [showAddGoal,setShowAddGoal] = useState(false);
  const [newGoal,setNewGoal]         = useState({
    kra:"attendance",goal_title:"",description:"",
    target_value:100,current_value:0,unit:"%",
    start_date:new Date().toISOString().split("T")[0],
    due_date:"",status:"not_started"
  });
  const [goalFilter,setGoalFilter]   = useState("all");
  const [currentTime,setCurrentTime] = useState(new Date());
  const [signedOut,setSignedOut]     = useState(false);
  const [signoutTime,setSignoutTime] = useState("");
  const [currentWeek]                = useState(getCurrentWeek());
  const [weeklyText,setWeeklyText]   = useState("");
  const [weeklySubmitted,setWeeklySubmitted] = useState(false);

  const [setupProfile,setSetupProfile] = useState({
    full_name:"",civil_id:"",phone_number:"",
    department:"",assigned_mentor:"",gpa:"",
    date_of_birth:"",gender:"",nationality:"Omani",
  });

  const excuseRef=useRef();
  const weeklyPhotoRef=useRef();

  const [profile,setProfile] = useState({
    full_name:"",civil_id:"",phone_number:"",
    department:"",assigned_mentor:"",gpa:"",
  });
  const [attendance,setAttendance] = useState({
    report_date:new Date().toISOString().split("T")[0],attended:false,
  });

  const s = {
    page:{minHeight:"100vh",background:HW.dark,color:HW.text,
      fontFamily:"sans-serif",padding:"16px 16px 100px 16px",
      maxWidth:"100%",margin:"0 auto"},
    card:{background:HW.surface,border:`1px solid ${HW.border}`,
      borderRadius:16,padding:16,marginBottom:16},
    input:{background:HW.surface2,border:`1px solid ${HW.border}`,
      color:HW.text,borderRadius:10,padding:"12px 14px",width:"100%",
      fontFamily:"inherit",fontSize:16,boxSizing:"border-box"},
    label:{fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",
      letterSpacing:".06em",display:"block",marginBottom:6},
    btn:{padding:"12px 20px",borderRadius:10,border:"none",
      fontWeight:700,cursor:"pointer",fontSize:14},
    th:{textAlign:"left",padding:"8px 10px",fontSize:11,fontWeight:700,
      color:HW.muted,textTransform:"uppercase",borderBottom:`1px solid ${HW.border}`},
    td:{padding:"10px 10px",fontSize:13,borderBottom:`1px solid rgba(51,51,51,.6)`},
  };

  const statusColors={
    active:{bg:`${HW.red}20`,color:HW.red},
    inactive:{bg:"rgba(136,136,136,.15)",color:HW.muted},
    transferred:{bg:"rgba(255,165,0,.15)",color:"#FFA500"},
    dropped:{bg:"rgba(136,136,136,.2)",color:"#666"},
  };

  const eventIcons={
    joined:"🟢",dropped:"🔴",transferred:"🔄",
    mentor_changed:"👤",dept_changed:"🏢",
    reactivated:"✅",note:"📝",inactive:"⏸"
  };

  function kpiColor(score){return score>=80?HW.red:score>=60?"#FFA500":"#666";}

  useEffect(()=>{
    const interval=setInterval(()=>{
      const now=new Date();
      setCurrentTime(now);
      const t=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
      setIsLate(t>MAX_SIGNIN);
    },1000);
    return ()=>clearInterval(interval);
  },[]);

  useEffect(()=>{
    const check=()=>{
      const now=new Date();
      if(now.getHours()===16&&now.getMinutes()===30){
        setShowReminder(true);
        if(Notification.permission==="granted")
          new Notification("Huawei TechTrack",{body:"Reminder: Submit your weekly report!"});
      }
    };
    if(Notification.permission==="default") Notification.requestPermission();
    const interval=setInterval(check,60000);
    return ()=>clearInterval(interval);
  },[]);

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>{
      if(session) handleSession(session.user);
    });
    const{data:listener}=supabase.auth.onAuthStateChange((_e,session)=>{
      if(session) handleSession(session.user);
      else{setUser(null);setView("login");}
    });
    return ()=>listener.subscription.unsubscribe();
  },[]);

  useEffect(()=>{
    if(view!=="mgmt") return;
    fetchTodaySignins();
    const channel=supabase.channel("live-signins")
      .on("postgres_changes",{event:"INSERT",schema:"public",table:"daily_reports"},
        async(payload)=>{
          const r=payload.new;
          const{data:t}=await supabase.from("trainees")
            .select("full_name,department").eq("id",r.trainee_id).single();
          setLiveSignins(prev=>[{
            id:r.id,full_name:t?.full_name||"Unknown",
            department:t?.department||"—",signin_time:r.signin_time,
            attended:r.attended,report_date:r.report_date,
            penalty_applied:r.penalty_applied,timestamp:new Date(),
          },...prev].slice(0,50));
        }).subscribe();
    return ()=>supabase.removeChannel(channel);
  },[view]);

  async function fetchTodaySignins(){
    const today=new Date().toISOString().split("T")[0];
    const{data}=await supabase.from("daily_reports")
      .select("*,trainees(full_name,department)")
      .eq("report_date",today).order("created_at",{ascending:false});
    if(data){
      setLiveSignins(data.map(r=>({
        id:r.id,full_name:r.trainees?.full_name||"Unknown",
        department:r.trainees?.department||"—",
        signin_time:r.signin_time,attended:r.attended,
        report_date:r.report_date,penalty_applied:r.penalty_applied,
        timestamp:new Date(r.created_at),
      })));
    }
  }

  async function handleSession(authUser){
    setUser(authUser);
    const{data}=await supabase.from("profiles")
      .select("role,trainee_id,profile_completed").eq("id",authUser.id).single();
    if(data){
      if(data.role==="management"){
        setView("mgmt");fetchTrainees();fetchOkrs();fetchAllData();
      } else {
        if(!data.profile_completed){setView("setup");}
        else {
          setView("trainee");
          if(data.trainee_id){
            setTraineeId(data.trainee_id);
            const{data:t}=await supabase.from("trainees")
              .select("*").eq("id",data.trainee_id).single();
            if(t){
              setProfile({full_name:t.full_name||"",civil_id:t.civil_id||"",
                phone_number:t.phone_number||"",department:t.department||"",
                assigned_mentor:t.assigned_mentor||"",gpa:t.gpa||""});
              setTraineeName(t.full_name.split(" ")[0]);
              setTimeout(()=>setShowGreeting(true),500);
            }
            fetchGoals(data.trainee_id);
            fetchWeeklyReport(data.trainee_id);
            const today=new Date().toISOString().split("T")[0];
            const{data:todayReport}=await supabase.from("daily_reports")
              .select("signout_time").eq("trainee_id",data.trainee_id)
              .eq("report_date",today).single();
            if(todayReport?.signout_time){setSignedOut(true);setSignoutTime(todayReport.signout_time);}
            const start=new Date();start.setDate(1);
            const{data:tc}=await supabase.from("traffic_excuses")
              .select("id").eq("trainee_id",data.trainee_id)
              .gte("created_at",start.toISOString());
            if(tc) setTrafficCount(tc.length);
          }
        }
      }
    }
  }

  async function fetchWeeklyReport(tid){
    const week=getCurrentWeek();
    const{data}=await supabase.from("daily_reports")
      .select("*").eq("trainee_id",tid)
      .eq("week_start",week.week_start).single();
    if(data){
      setWeeklySubmitted(true);setWeeklyText(data.weekly_tasks||"");
      if(data.kpi_score) setAiResult({
        kpi_score:data.kpi_score,
        pie_chart:data.pie_chart_json?JSON.parse(data.pie_chart_json):null,
        talent_notes:data.talent_notes,summary:data.report_text,
      });
    }
  }

  async function saveSetupProfile(){
    if(!setupProfile.full_name||!setupProfile.civil_id){
      setMsg("Please fill Full Name and Civil ID.");return;
    }
    setLoading(true);setMsg("");
    const{data:traineeData,error:tErr}=await supabase.from("trainees").upsert({
      full_name:setupProfile.full_name,civil_id:setupProfile.civil_id,
      phone_number:setupProfile.phone_number,department:setupProfile.department,
      assigned_mentor:setupProfile.assigned_mentor,
      gpa:setupProfile.gpa?parseFloat(setupProfile.gpa):null,
      date_of_birth:setupProfile.date_of_birth||null,
      gender:setupProfile.gender||null,nationality:setupProfile.nationality||null,
      joining_date:new Date().toISOString().split("T")[0],
    },{onConflict:"civil_id"}).select().single();
    if(tErr){setMsg("Error: "+tErr.message);setLoading(false);return;}
    await supabase.from("profiles").update({
      trainee_id:traineeData.id,profile_completed:true,
    }).eq("id",user.id);
    setTraineeId(traineeData.id);
    setProfile({full_name:setupProfile.full_name,civil_id:setupProfile.civil_id,
      phone_number:setupProfile.phone_number,department:setupProfile.department,
      assigned_mentor:setupProfile.assigned_mentor,gpa:setupProfile.gpa});
    setTraineeName(setupProfile.full_name.split(" ")[0]);
    fetchGoals(traineeData.id);
    setLoading(false);setView("trainee");
    setTimeout(()=>setShowGreeting(true),500);
  }

  async function login(){
    setLoading(true);setMsg("");
    const{error}=await supabase.auth.signInWithPassword({email,password});
    if(error) setMsg(error.message);
    setLoading(false);
  }

  async function signup(){
    if(!email||!password){setMsg("Please enter email and password.");return;}
    if(password!==confirmPwd){setMsg("Passwords do not match.");return;}
    if(password.length<6){setMsg("Password must be at least 6 characters.");return;}
    setLoading(true);setMsg("");
    const{data:authData,error}=await supabase.auth.signUp({email,password});
    if(error){setMsg(error.message);setLoading(false);return;}
    await supabase.from("profiles").insert({
      id:authData.user.id,email,role:"trainee",profile_completed:false});
    setMsg("✅ Account created! You can now log in.");
    setLoading(false);
    setTimeout(()=>{setView("login");setMsg("");setConfirmPwd("");},2000);
  }

  async function logout(){
    await supabase.auth.signOut();
    setView("login");setUser(null);setSelected(null);setMsg("");
    setAiResult(null);setWeeklyPhotoFile(null);setWeeklyPhotoPreview(null);
    setLocationOk(false);setGeoStatus("idle");setGeoMsg("");
    setTraineeId(null);setGoals([]);setLiveSignins([]);
    setShowGreeting(false);setTraineeName("");
    setSignedOut(false);setSignoutTime("");
    setWeeklySubmitted(false);setWeeklyText("");
  }

  function checkLocation(){
    setGeoStatus("checking");setGeoMsg("📍 Checking your location…");
    if(!navigator.geolocation){
      setGeoStatus("error");setGeoMsg("❌ Browser does not support location.");return;
    }
    navigator.geolocation.getCurrentPosition(
      (pos)=>{
        const dist=getDistance(pos.coords.latitude,pos.coords.longitude,WORK_LAT,WORK_LNG);
        if(dist<=WORK_RADIUS){
          setLocationOk(true);setGeoStatus("ok");
          setGeoMsg(`✅ Verified — ${Math.round(dist)}m from workplace.`);
        } else {
          setLocationOk(false);setGeoStatus("error");
          setGeoMsg(`❌ ${Math.round(dist)}m away. Must be within ${WORK_RADIUS}m.`);
        }
      },
      ()=>{setGeoStatus("error");setGeoMsg("❌ Location denied.");},
      {enableHighAccuracy:true,timeout:10000}
    );
  }

  function handleExcusePhoto(e){const f=e.target.files[0];if(!f)return;setExcusePhoto(f);setExcusePreview(URL.createObjectURL(f));}
  function handleWeeklyPhoto(e){const f=e.target.files[0];if(!f)return;setWeeklyPhotoFile(f);setWeeklyPhotoPreview(URL.createObjectURL(f));}

  async function uploadFile(bucket,path,file){
    const{error}=await supabase.storage.from(bucket).upload(path,file);
    if(error) return null;
    return supabase.storage.from(bucket).getPublicUrl(path).data.publicUrl;
  }

  async function submitAttendance(){
    if(!locationOk){setMsg("📍 Please verify your location first.");return;}
    const now=new Date();
    const timeStr=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    if(timeStr>MAX_SIGNIN&&!excuseText){
      setMsg("⚠️ Past 9:00 AM — please provide an excuse.");return;
    }
    if(excuseType==="traffic"){
      if(trafficCount>=2){setMsg("❌ Traffic excuse limit reached.");return;}
      if(!excusePhoto){setMsg("⚠️ Traffic excuse requires proof photo.");return;}
    }
    const penaltyApplied=!attendance.attended||(timeStr>MAX_SIGNIN&&!excuseText);
    setLoading(true);setMsg("");
    const tid=traineeId;const ts=Date.now();
    let excusePhotoUrl=null;
    if(excusePhoto)
      excusePhotoUrl=await uploadFile("report-photos",`${tid}/excuse_${ts}.jpg`,excusePhoto);
    if(excuseType==="traffic"&&excusePhotoUrl){
      await supabase.from("traffic_excuses").insert({
        trainee_id:tid,report_date:attendance.report_date,photo_url:excusePhotoUrl});
      setTrafficCount(c=>c+1);
    }
    if(penaltyApplied){
      await supabase.from("penalties").insert({
        trainee_id:tid,report_date:attendance.report_date,
        reason:!attendance.attended?"Absent — no attendance recorded":"Late sign-in after 9:00 AM without excuse",
        amount:PENALTY_PCT,
      });
    }
    await supabase.from("daily_reports").upsert({
      trainee_id:tid,report_date:attendance.report_date,
      attended:attendance.attended,signin_time:timeStr,
      excuse_type:excuseType||null,excuse_text:excuseText||null,
      excuse_photo_url:excusePhotoUrl,traffic_excuse:excuseType==="traffic",
      penalty_applied:penaltyApplied,penalty_amount:penaltyApplied?PENALTY_PCT:0,
    },{onConflict:"trainee_id,report_date"});
    setMsg(penaltyApplied
      ?`✅ Attendance recorded — ⚠️ Penalty ${PENALTY_PCT}% applied.`
      :"✅ Attendance recorded successfully!");
    setLoading(false);
  }

  async function submitSignOut(){
    if(!traineeId){setMsg("Please submit attendance first.");return;}
    setLoading(true);setMsg("");
    const now=new Date();
    const timeStr=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    const today=new Date().toISOString().split("T")[0];
    const{error}=await supabase.from("daily_reports")
      .update({signout_time:timeStr}).eq("trainee_id",traineeId).eq("report_date",today);
    if(error){setMsg("Error: "+error.message);}
    else{setSignedOut(true);setSignoutTime(timeStr);setMsg(`✅ Signed out at ${timeStr}`);}
    setLoading(false);
  }

  async function submitWeeklyReport(){
    if(!weeklyText){setMsg("Please write your weekly tasks.");return;}
    setLoading(true);setAiLoading(true);setMsg("");setAiResult(null);
    const tid=traineeId;
    if(!tid){setMsg("Please submit attendance first.");setLoading(false);setAiLoading(false);return;}
    const ts=Date.now();
    const week=getCurrentWeek();
    const photoUrl=weeklyPhotoFile
      ?await uploadFile("report-photos",`${tid}/weekly_${ts}.jpg`,weeklyPhotoFile):null;
    setMsg("🤖 AI analyzing your report…");
    let ai=null;
    try{ai=await analyzeReport(weeklyText);setAiResult(ai);}catch(e){console.error(e);}
    await supabase.from("daily_reports").upsert({
      trainee_id:tid,report_date:week.week_start,
      week_start:week.week_start,week_end:week.week_end,
      weekly_tasks:weeklyText,weekly_photo_url:photoUrl,
      report_text:ai?.summary||weeklyText.substring(0,100),
      kpi_score:ai?.kpi_score||null,
      pie_chart_json:ai?.pie_chart?JSON.stringify(ai.pie_chart):null,
      talent_notes:ai?.talent_notes||null,
    },{onConflict:"trainee_id,report_date"});
    setWeeklySubmitted(true);
    setMsg("✅ Weekly report submitted!");
    setLoading(false);setAiLoading(false);
  }

  async function fetchGoals(tid){
    const{data}=await supabase.from("goals").select("*")
      .eq("trainee_id",tid).order("created_at",{ascending:false});
    if(data) setGoals(data);
  }

  async function addGoal(){
    if(!newGoal.goal_title){setMsg("Please enter a goal title.");return;}
    if(!traineeId){setMsg("Please complete your profile first.");return;}
    setLoading(true);setMsg("");
    const{error}=await supabase.from("goals").insert({...newGoal,trainee_id:traineeId});
    if(error){setMsg(error.message);}
    else{
      setMsg("✅ Goal added!");setShowAddGoal(false);
      setNewGoal({kra:"attendance",goal_title:"",description:"",
        target_value:100,current_value:0,unit:"%",
        start_date:new Date().toISOString().split("T")[0],
        due_date:"",status:"not_started"});
      fetchGoals(traineeId);
    }
    setLoading(false);
  }

  async function updateGoal(id,newValue,newStatus){
    const updates={current_value:newValue};
    if(newStatus) updates.status=newStatus;
    else if(newValue>=goals.find(g=>g.id===id)?.target_value) updates.status="completed";
    await supabase.from("goals").update(updates).eq("id",id);
    fetchGoals(traineeId);setMsg("✅ Goal updated!");
  }

  async function deleteGoal(id){
    await supabase.from("goals").delete().eq("id",id);
    fetchGoals(traineeId);setMsg("✅ Goal deleted.");
  }

  async function fetchSelectedGoals(tid){
    const{data}=await supabase.from("goals").select("*")
      .eq("trainee_id",tid).order("created_at",{ascending:false});
    if(data) setSelectedGoals(data);
  }

  async function fetchTrainees(){
    setLoading(true);
    const{data}=await supabase.from("trainees").select("*").order("full_name");
    if(data) setTrainees(data);setLoading(false);
  }

  async function fetchOkrs(){
    const{data}=await supabase.from("okrs").select("*").order("department");
    if(data) setOkrs(data);
  }

  async function fetchAllData(){
    const{data:reps}=await supabase.from("daily_reports").select("*");
    const{data:pens}=await supabase.from("penalties").select("*");
    if(reps) setAllReports(reps);if(pens) setAllPenalties(pens);
  }

  async function fetchReports(tid){
    const{data}=await supabase.from("daily_reports").select("*")
      .eq("trainee_id",tid).order("report_date",{ascending:false});
    if(data) setReports(data);
  }

  async function fetchLogs(tid){
    const{data}=await supabase.from("trainee_logs").select("*")
      .eq("trainee_id",tid).order("created_at",{ascending:false});
    if(data) setLogs(data);
  }

  async function fetchPenalties(tid){
    const{data}=await supabase.from("penalties").select("*")
      .eq("trainee_id",tid).order("created_at",{ascending:false});
    if(data) setPenalties(data);
  }

  async function openProfile(t){
    setSelected({...t});setProfileTab("timeline");setMsg("");
    await fetchReports(t.id);await fetchLogs(t.id);
    await fetchPenalties(t.id);await fetchSelectedGoals(t.id);
  }

  async function logEvent(tid,type,desc,old="",nw=""){
    await supabase.from("trainee_logs").insert({
      trainee_id:tid,event_type:type,description:desc,
      old_value:old,new_value:nw,logged_by:user?.email||"manager"
    });
  }

  async function saveProfile(){
    const changes=[];
    if(selected.department!==selected._original?.department)
      changes.push(logEvent(selected.id,"dept_changed",
        `Department changed from ${selected._original?.department} to ${selected.department}`,
        selected._original?.department,selected.department));
    if(selected.assigned_mentor!==selected._original?.assigned_mentor)
      changes.push(logEvent(selected.id,"mentor_changed",
        `Mentor changed from ${selected._original?.assigned_mentor} to ${selected.assigned_mentor}`,
        selected._original?.assigned_mentor,selected.assigned_mentor));
    const{error}=await supabase.from("trainees").update({
      department:selected.department,assigned_mentor:selected.assigned_mentor,gpa:selected.gpa
    }).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await Promise.all(changes);
    setMsg("✅ Saved!");fetchTrainees();fetchLogs(selected.id);
  }

  async function changeStatus(newStatus){
    const messages={
      dropped:"Trainee has left/dropped the program",
      inactive:"Trainee marked as inactive",
      transferred:"Trainee transferred to another department",
      active:"Trainee reactivated in the program",
    };
    const updateData={status:newStatus};
    if(newStatus==="dropped") updateData.quitting_date=new Date().toISOString().split("T")[0];
    if(newStatus==="active") updateData.quitting_date=null;
    const{error}=await supabase.from("trainees").update(updateData).eq("id",selected.id);
    if(error){setMsg(error.message);return;}
    await logEvent(selected.id,newStatus==="active"?"reactivated":newStatus,
      messages[newStatus],selected.status,newStatus);
    setSelected({...selected,status:newStatus,...updateData});
    fetchTrainees();fetchLogs(selected.id);
    setMsg(`✅ Status updated to ${newStatus}`);
  }

  async function updateOkr(id,newValue){
    await supabase.from("okrs").update({current:newValue}).eq("id",id);fetchOkrs();
  }

  async function addOkr(){
    if(!newOkr.department||!newOkr.objective||!newOkr.key_result){
      setMsg("Please fill all OKR fields.");return;
    }
    await supabase.from("okrs").insert({...newOkr,created_by:user?.email});
    setShowAddOkr(false);
    setNewOkr({department:"",objective:"",key_result:"",target:100,current:0,unit:"%",due_date:""});
    fetchOkrs();setMsg("✅ OKR added!");
  }

  async function exportExcel(){
    setMsg("📊 Preparing export…");
    const{data:allT}=await supabase.from("trainees").select("*").order("full_name");
    const{data:allR}=await supabase.from("daily_reports").select("*").order("report_date");
    const{data:allP}=await supabase.from("penalties").select("*").order("created_at");
    const{data:allG}=await supabase.from("goals").select("*").order("created_at");
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allT||[]).map(t=>({
        "Full Name":t.full_name,"Civil ID":t.civil_id,"Phone":t.phone_number,
        "Department":t.department,"Mentor":t.assigned_mentor,"GPA":t.gpa,
        "Gender":t.gender||"—","Nationality":t.nationality||"—",
        "Status":t.status,"Joining Date":t.joining_date,
      }))),"Trainees");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allR||[]).filter(r=>!r.week_start).map(r=>{
        const trainee=(allT||[]).find(t=>t.id===r.trainee_id);
        const dateObj=new Date(r.report_date+"T00:00:00");
        return{
          "Full Name":trainee?.full_name||"—","Department":trainee?.department||"—",
          "Day":dateObj.toLocaleDateString("en-GB",{weekday:"long"}),
          "Date":dateObj.toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"}),
          "Attended":r.attended?"Yes":"No","Time In":r.signin_time||"—",
          "Time Out":r.signout_time||"—","Penalty":r.penalty_applied?`-${r.penalty_amount}%`:"None",
        };
      })),"Daily Attendance");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allR||[]).filter(r=>r.weekly_tasks).map(r=>{
        const trainee=(allT||[]).find(t=>t.id===r.trainee_id);
        return{
          "Full Name":trainee?.full_name||"—","Department":trainee?.department||"—",
          "Week":r.week_start+" to "+r.week_end,"Weekly Tasks":r.weekly_tasks||"—",
          "KPI Score":r.kpi_score||"—","AI Summary":r.report_text||"—",
          "Talent Notes":r.talent_notes||"—",
        };
      })),"Weekly Reports");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allP||[]).map(p=>({
        "Full Name":(allT||[]).find(t=>t.id===p.trainee_id)?.full_name||"—",
        "Date":p.report_date,"Reason":p.reason,"Deduction":`-${p.amount}%`,
      }))),"Penalties");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allG||[]).map(g=>({
        "Full Name":(allT||[]).find(t=>t.id===g.trainee_id)?.full_name||"—",
        "KRA":g.kra,"Goal":g.goal_title,"Target":g.target_value,
        "Current":g.current_value,"Unit":g.unit,"Status":g.status,
        "Progress":Math.min((g.current_value/g.target_value)*100,100).toFixed(0)+"%",
      }))),"Goals");
    XLSX.writeFile(wb,`Huawei TechTrack_${new Date().toISOString().split("T")[0]}.xlsx`);
    setMsg("✅ Excel exported!");
  }

  async function exportPDF(trainee){
    const doc=new jsPDF();const t=trainee||selected;
    doc.setFillColor(207,10,44);doc.rect(0,0,210,40,"F");
    doc.setTextColor(255,255,255);doc.setFontSize(18);doc.setFont("helvetica","bold");
    doc.text("Huawei TechTrack — Performance Report",14,18);
    doc.setFontSize(10);doc.setFont("helvetica","normal");
    doc.text(`Generated: ${new Date().toLocaleDateString("en-GB")}`,14,30);
    doc.setTextColor(0,0,0);doc.setFontSize(13);doc.setFont("helvetica","bold");
    doc.text("Trainee Information",14,52);
    autoTable(doc,{startY:56,
      head:[["Field","Details"]],
      body:[["Full Name",t.full_name||"—"],["Civil ID",t.civil_id||"—"],
        ["Department",t.department||"—"],["Mentor",t.assigned_mentor||"—"],
        ["Status",t.status||"—"],["Joining Date",t.joining_date||"—"]],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    const penY=doc.lastAutoTable.finalY+10;
    doc.setFontSize(13);doc.setFont("helvetica","bold");doc.text("Penalties",14,penY);
    const{data:pens}=await supabase.from("penalties").select("*").eq("trainee_id",t.id);
    autoTable(doc,{startY:penY+4,
      head:[["Date","Reason","Deduction"]],
      body:pens?.length>0?pens.map(p=>[p.report_date,p.reason,`-${p.amount}%`])
        :[["—","No penalties","—"]],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    doc.setTextColor(150,150,150);doc.setFontSize(9);doc.setFont("helvetica","normal");
    doc.text("Huawei TechTrack • Powered by Arjwan Sabir • Confidential",14,285);
    doc.save(`${t.full_name}_report.pdf`);
  }

  function getAnalytics(){
    const active=trainees.filter(t=>t.status==="active");
    const totalReports=allReports.length;
    const attended=allReports.filter(r=>r.attended).length;
    const attendanceRate=totalReports>0?((attended/totalReports)*100).toFixed(1):0;
    const kpiReports=allReports.filter(r=>r.kpi_score);
    const avgKpi=kpiReports.length>0
      ?(kpiReports.reduce((a,r)=>a+r.kpi_score,0)/kpiReports.length).toFixed(1):0;
    const traineeKpi=trainees.map(t=>{
      const tr=allReports.filter(r=>r.trainee_id===t.id&&r.kpi_score);
      const avg=tr.length>0?tr.reduce((a,r)=>a+r.kpi_score,0)/tr.length:0;
      const tp=allPenalties.filter(p=>p.trainee_id===t.id).length;
      return{...t,avgKpi:avg.toFixed(1),penalties:tp,
        reports:allReports.filter(r=>r.trainee_id===t.id).length};
    }).sort((a,b)=>b.avgKpi-a.avgKpi);
    const depts=[...new Set(trainees.map(t=>t.department).filter(Boolean))];
    const deptData=depts.map(dept=>{
      const dT=trainees.filter(t=>t.department===dept);
      const dR=allReports.filter(r=>dT.some(t=>t.id===r.trainee_id)&&r.kpi_score);
      const avg=dR.length>0?dR.reduce((a,r)=>a+r.kpi_score,0)/dR.length:0;
      return{label:dept.substring(0,8),value:Math.round(avg)};
    });
    const atRisk=traineeKpi.filter(t=>parseFloat(t.avgKpi)<60||t.penalties>2);
    return{active,attendanceRate,avgKpi,totalPenalties:allPenalties.length,
      traineeKpi,deptData,atRisk};
  }

  const filteredGoals=goalFilter==="all"?goals
    :goals.filter(g=>g.kra===goalFilter||g.status===goalFilter);

  const hh=String(currentTime.getHours()).padStart(2,"0");
  const mm=String(currentTime.getMinutes()).padStart(2,"0");
  const ss=String(currentTime.getSeconds()).padStart(2,"0");
  const clockColor=isLate?HW.red:"#34d399";

  const TraineeNav = () => (
    <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
      width:"100%",maxWidth:480,background:HW.surface,
      borderTop:`1px solid ${HW.border}`,
      display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)"}}>
      {[
        {id:"attendance",icon:"✅",label:"Attend"},
        {id:"weekly",icon:"📅",label:"Weekly"},
        {id:"goals",icon:"🎯",label:"Goals"},
      ].map(tab=>(
        <button key={tab.id} onClick={()=>setTraineeTab(tab.id)}
          style={{flex:1,padding:"12px 4px 8px",border:"none",
            background:"none",cursor:"pointer",
            borderTop:traineeTab===tab.id?`3px solid ${HW.red}`:"3px solid transparent",
            display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
          <span style={{fontSize:22}}>{tab.icon}</span>
          <span style={{fontSize:10,fontWeight:700,
            color:traineeTab===tab.id?HW.red:HW.muted}}>{tab.label}</span>
        </button>
      ))}
    </div>
  );

  const MgmtNav = () => (
    <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
      width:"100%",maxWidth:480,background:HW.surface,
      borderTop:`1px solid ${HW.border}`,
      display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)"}}>
      {["trainees","live","analytics","okr"].map(tab=>(
        <button key={tab} onClick={()=>{setMgmtTab(tab);setSelected(null);setMsg("");}}
          style={{flex:1,padding:"12px 4px 8px",border:"none",
            background:"none",cursor:"pointer",
            borderTop:mgmtTab===tab?`3px solid ${HW.red}`:"3px solid transparent",
            display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
          <span style={{fontSize:20}}>
            {tab==="trainees"?"👥":tab==="live"?"📡":tab==="analytics"?"📊":"🎯"}
          </span>
          <span style={{fontSize:10,fontWeight:700,
            color:mgmtTab===tab?HW.red:HW.muted}}>
            {tab==="trainees"?"Trainees":tab==="live"?"Live":tab==="analytics"?"Analytics":"OKR"}
          </span>
        </button>
      ))}
    </div>
  );

  // ══ LOGIN ══
  if(view==="login") return (
    <div style={{...s.page,display:"flex",alignItems:"center",justifyContent:"center",
      minHeight:"100vh",padding:20}}>
      <div style={{width:"100%",maxWidth:400}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <HuaweiLogo size={64}/>
          <h1 style={{margin:"16px 0 4px",fontSize:28,color:HW.text}}>Huawei TechTrack</h1>
          <p style={{color:HW.red,fontSize:12,fontWeight:700,
            letterSpacing:".1em",textTransform:"uppercase"}}>Powered by Arjwan Sabir</p>
          <div style={{fontSize:11,color:HW.muted,marginTop:6}}>
            🎨 {randomTheme.name} Theme
          </div>
        </div>
        <label style={s.label}>Email</label>
        <input style={{...s.input,marginBottom:16}} value={email}
          onChange={e=>setEmail(e.target.value)} placeholder="you@example.com"
          type="email" autoComplete="email"/>
        <label style={s.label}>Password</label>
        <input style={{...s.input,marginBottom:24}} type="password"
          value={password} onChange={e=>setPassword(e.target.value)}
          placeholder="Password" autoComplete="current-password"
          onKeyDown={e=>e.key==="Enter"&&login()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,
          width:"100%",padding:16,fontSize:17,opacity:loading?0.6:1}}
          onClick={login} disabled={loading}>
          {loading?"Signing in…":"Sign In →"}
        </button>
        <div style={{display:"flex",alignItems:"center",gap:10,margin:"20px 0"}}>
          <div style={{flex:1,height:1,background:HW.border}}/>
          <span style={{color:HW.muted,fontSize:12}}>OR</span>
          <div style={{flex:1,height:1,background:HW.border}}/>
        </div>
        <button style={{...s.btn,background:HW.surface2,color:HW.text,
          width:"100%",padding:16,border:`1px solid ${HW.border}`,fontSize:15}}
          onClick={()=>{setView("signup");setMsg("");}}>
          Create Trainee Account
        </button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
          fontSize:14,marginTop:16,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══ SIGNUP ══
  if(view==="signup") return (
    <div style={{...s.page,display:"flex",alignItems:"center",justifyContent:"center",
      minHeight:"100vh",padding:20}}>
      <div style={{width:"100%",maxWidth:400}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <HuaweiLogo size={48}/>
          <h2 style={{margin:"12px 0 4px",color:HW.text}}>Create Account</h2>
          <p style={{color:HW.muted,fontSize:14}}>Register as a new trainee</p>
        </div>
        <label style={s.label}>Email Address</label>
        <input style={{...s.input,marginBottom:14}} value={email}
          onChange={e=>setEmail(e.target.value)} placeholder="yourname@example.com"
          type="email" autoComplete="email"/>
        <label style={s.label}>Password</label>
        <input style={{...s.input,marginBottom:14}} type="password"
          value={password} onChange={e=>setPassword(e.target.value)}
          placeholder="Minimum 6 characters"/>
        <label style={s.label}>Confirm Password</label>
        <input style={{...s.input,marginBottom:24}} type="password"
          value={confirmPwd} onChange={e=>setConfirmPwd(e.target.value)}
          placeholder="Repeat your password"
          onKeyDown={e=>e.key==="Enter"&&signup()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,
          width:"100%",padding:16,fontSize:16,opacity:loading?0.6:1}}
          onClick={signup} disabled={loading}>
          {loading?"Creating account…":"Create Account →"}
        </button>
        <button style={{...s.btn,background:"none",color:HW.muted,
          width:"100%",padding:14,marginTop:8,fontSize:15}}
          onClick={()=>{setView("login");setMsg("");}}>← Back to Login</button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
          fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══ PROFILE SETUP ══
  if(view==="setup") return (
    <div style={{...s.page,paddingBottom:32}}>
      <div style={{textAlign:"center",marginBottom:24,paddingTop:20}}>
        <HuaweiLogo size={52}/>
        <h2 style={{margin:"12px 0 6px",fontSize:22,color:HW.text}}>Welcome! 🎉</h2>
        <p style={{color:HW.muted,fontSize:13}}>Complete your profile once to get started.</p>
      </div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>👤 Personal Info</h3>
        <label style={s.label}>Full Name *</label>
        <input style={{...s.input,marginBottom:12}} placeholder="Ahmed Mohammed Al-Rashidi"
          value={setupProfile.full_name}
          onChange={e=>setSetupProfile({...setupProfile,full_name:e.target.value})}/>
        <label style={s.label}>Civil ID *</label>
        <input style={{...s.input,marginBottom:12}} placeholder="10234567"
          value={setupProfile.civil_id}
          onChange={e=>setSetupProfile({...setupProfile,civil_id:e.target.value})}/>
        <label style={s.label}>Date of Birth</label>
        <input style={{...s.input,marginBottom:12}} type="date"
          value={setupProfile.date_of_birth}
          onChange={e=>setSetupProfile({...setupProfile,date_of_birth:e.target.value})}/>
        <label style={s.label}>Gender</label>
        <select style={{...s.input,marginBottom:12}} value={setupProfile.gender}
          onChange={e=>setSetupProfile({...setupProfile,gender:e.target.value})}>
          <option value="">Select…</option>
          <option value="Male">Male / ذكر</option>
          <option value="Female">Female / أنثى</option>
        </select>
        <label style={s.label}>Nationality</label>
        <input style={{...s.input}} placeholder="e.g. Omani"
          value={setupProfile.nationality}
          onChange={e=>setSetupProfile({...setupProfile,nationality:e.target.value})}/>
      </div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>📚 Academic Info</h3>
        <label style={s.label}>Department</label>
        <input style={{...s.input,marginBottom:12}} placeholder="e.g. Engineering"
          value={setupProfile.department}
          onChange={e=>setSetupProfile({...setupProfile,department:e.target.value})}/>
        <label style={s.label}>GPA</label>
        <input style={{...s.input,marginBottom:12}} type="number"
          placeholder="e.g. 3.85" min="0" max="4" step="0.01"
          value={setupProfile.gpa}
          onChange={e=>setSetupProfile({...setupProfile,gpa:e.target.value})}/>
        <label style={s.label}>Assigned Mentor</label>
        <input style={s.input} placeholder="e.g. Dr. Fatima"
          value={setupProfile.assigned_mentor}
          onChange={e=>setSetupProfile({...setupProfile,assigned_mentor:e.target.value})}/>
      </div>
      <div style={s.card}>
        <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>📞 Contact</h3>
        <label style={s.label}>Phone Number</label>
        <input style={s.input} placeholder="+968-9100-0001"
          value={setupProfile.phone_number} type="tel"
          onChange={e=>setSetupProfile({...setupProfile,phone_number:e.target.value})}/>
      </div>
      <button style={{...s.btn,background:HW.red,color:HW.white,
        width:"100%",padding:16,fontSize:16,opacity:loading?0.6:1}}
        onClick={saveSetupProfile} disabled={loading}>
        {loading?"Saving…":"Complete Profile & Start →"}
      </button>
      {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
        fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
    </div>
  );

  // ══ TRAINEE DASHBOARD ══
  if(view==="trainee") return (
    <div style={s.page}>
      {showReminder&&<ReminderPopup onDismiss={()=>setShowReminder(false)}/>}
      {showGreeting&&<GreetingPopup name={traineeName} onDismiss={()=>setShowGreeting(false)}/>}
      <TraineeNav/>

      {/* Topbar */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
        marginBottom:12,paddingTop:4}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <HuaweiLogo size={28}/>
          <div>
            <div style={{fontWeight:700,fontSize:15,color:HW.text}}>
              {traineeName?`Hi, ${traineeName}!`:"Huawei TechTrack"}
            </div>
            <div style={{fontSize:10,color:HW.muted}}>{user?.email}</div>
          </div>
        </div>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:20,fontWeight:800,color:clockColor,fontFamily:"monospace"}}>
            {hh}:{mm}:{ss}
          </div>
          <div style={{fontSize:10,color:clockColor,fontWeight:700}}>
            {isLate?"⚠️ Past 9 AM":"✅ On Time"}
          </div>
        </div>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:4}}>
        <button style={{...s.btn,background:HW.surface2,color:HW.muted,
          fontSize:11,padding:"6px 10px"}}
          onClick={()=>setShowGreeting(true)}>👋 Greeting</button>
        <button style={{...s.btn,background:HW.surface2,color:HW.muted,
          fontSize:11,padding:"6px 10px"}}
          onClick={()=>setShowReminder(true)}>⏰ Reminder</button>
        <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,
          fontSize:11,padding:"6px 10px",marginLeft:"auto"}}
          onClick={logout}>Sign out</button>
      </div>
      <div style={{height:1,background:HW.border,margin:"12px 0 16px"}}/>

      {/* ══ ATTENDANCE TAB ══ */}
      {traineeTab==="attendance"&&(
        <div>
          <div style={{...s.card,textAlign:"center",border:`1px solid ${clockColor}40`,
            background:`linear-gradient(135deg,${HW.surface},${HW.red}08)`}}>
            <div style={{fontSize:48,fontWeight:800,color:clockColor,
              fontFamily:"monospace",letterSpacing:4,
              textShadow:`0 0 30px ${clockColor}40`}}>
              {hh}:{mm}:{ss}
            </div>
            <div style={{fontSize:12,color:HW.muted,marginTop:4}}>
              {currentTime.toLocaleDateString("en-GB",{
                weekday:"long",day:"numeric",month:"long"})}
            </div>
            <div style={{marginTop:10,display:"inline-block",
              padding:"5px 16px",borderRadius:20,fontWeight:700,fontSize:12,
              background:`${clockColor}15`,color:clockColor,
              border:`1px solid ${clockColor}40`}}>
              {isLate?"⚠️ Past 9:00 AM — Excuse Required":"✅ On Time — Ready to Sign In"}
            </div>
          </div>

          {isLate&&(
            <div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}40`,
              borderRadius:12,padding:14,marginBottom:16,display:"flex",gap:10}}>
              <span style={{fontSize:24}}>🔒</span>
              <div>
                <div style={{fontWeight:700,color:HW.red,fontSize:14}}>Time Locked</div>
                <div style={{fontSize:12,color:HW.muted,marginTop:2}}>
                  Deadline was 9:00 AM. <b style={{color:HW.red}}>8.33% penalty</b> will apply.
                </div>
              </div>
            </div>
          )}

          <div style={s.card}>
            <h3 style={{marginBottom:14,color:HW.red,fontSize:16}}>✅ Sign Attendance</h3>
            <label style={s.label}>Date</label>
            <input style={{...s.input,marginBottom:14}} type="date"
              value={attendance.report_date}
              onChange={e=>setAttendance({...attendance,report_date:e.target.value})}/>
            <label style={s.label}>Sign-in Time (auto)</label>
            <div style={{background:HW.surface2,border:`1px solid ${HW.border}`,
              borderRadius:10,padding:"12px 14px",fontSize:20,
              color:clockColor,fontWeight:800,fontFamily:"monospace",marginBottom:14}}>
              {hh}:{mm}:{ss}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14,
              background:HW.surface2,borderRadius:10,padding:14}}>
              <input type="checkbox" id="att" checked={attendance.attended}
                style={{width:22,height:22}}
                onChange={e=>setAttendance({...attendance,attended:e.target.checked})}/>
              <label htmlFor="att" style={{fontSize:15,fontWeight:600,
                cursor:"pointer",color:HW.text}}>
                I confirm attendance today
              </label>
            </div>
            {!attendance.attended&&(
              <div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}40`,
                borderRadius:10,padding:12,marginBottom:14,fontSize:13,color:HW.red}}>
                ⚠️ <b>Penalty Warning:</b> Not marked as attended — <b>8.33% deduction</b>
              </div>
            )}
            {isLate&&(
              <div style={{background:HW.surface2,borderRadius:12,padding:14,
                border:`1px solid ${HW.red}40`,marginBottom:14}}>
                <div style={{fontWeight:700,color:HW.red,marginBottom:12,fontSize:15}}>
                  ⚠️ Excuse Required
                </div>
                <label style={s.label}>Excuse Type</label>
                <select style={{...s.input,marginBottom:12}} value={excuseType}
                  onChange={e=>setExcuseType(e.target.value)}>
                  <option value="">Select reason…</option>
                  <option value="traffic">🚗 Road Traffic {trafficCount>=2?"(LIMIT REACHED)":""}</option>
                  <option value="medical">🏥 Medical Emergency</option>
                  <option value="family">👨‍👩‍👧 Family Emergency</option>
                  <option value="other">📋 Other</option>
                </select>
                <label style={s.label}>Description *</label>
                <textarea style={{...s.input,height:90,resize:"vertical",marginBottom:12}}
                  placeholder="Please describe your reason…"
                  value={excuseText} onChange={e=>setExcuseText(e.target.value)}/>
                <label style={s.label}>
                  Proof Photo {excuseType==="traffic"?"(Required)":"(Optional)"}
                </label>
                <div style={{display:"flex",gap:10,alignItems:"center"}}>
                  <button style={{...s.btn,background:HW.surface,color:HW.text,
                    border:`1px dashed ${HW.border}`,flex:1}}
                    onClick={()=>excuseRef.current.click()}>
                    📷 {excusePhoto?"Change Photo":"Upload Proof"}
                  </button>
                  <input ref={excuseRef} type="file" accept="image/*"
                    style={{display:"none"}} onChange={handleExcusePhoto}/>
                  {excusePreview&&(
                    <img src={excusePreview} alt="proof"
                      style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>
                  )}
                </div>
              </div>
            )}

            {/* Location */}
            <div style={{background:HW.surface2,borderRadius:12,padding:14,
              marginBottom:14,
              border:locationOk?`1px solid ${HW.red}50`:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:10,fontSize:15}}>
                📍 Location Verification
              </div>
              <p style={{fontSize:12,color:HW.muted,marginBottom:10}}>
                Must be within <b style={{color:HW.text}}>{WORK_RADIUS}m</b> of workplace.
              </p>
              <button style={{...s.btn,
                background:locationOk?`${HW.red}20`:HW.red,
                color:locationOk?HW.red:HW.white,
                width:"100%",opacity:geoStatus==="checking"?0.6:1}}
                onClick={checkLocation} disabled={geoStatus==="checking"}>
                {geoStatus==="checking"?"📍 Checking…":
                 locationOk?"✅ Location Verified":"📍 Verify My Location"}
              </button>
              {geoMsg&&<p style={{fontSize:12,marginTop:8,
                color:geoStatus==="ok"?"#34d399":HW.red}}>{geoMsg}</p>}
            </div>

            <button style={{...s.btn,
              background:locationOk?HW.red:HW.surface2,
              color:locationOk?HW.white:HW.muted,
              width:"100%",padding:16,fontSize:16,
              opacity:loading?0.6:1,cursor:locationOk?"pointer":"not-allowed"}}
              onClick={submitAttendance} disabled={loading||!locationOk}>
              {loading?"Saving…":"✅ Submit Attendance"}
            </button>

            {/* Sign Out */}
            <div style={{marginTop:12,background:HW.surface2,borderRadius:12,padding:14,
              border:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:"#4f8ef7",marginBottom:10,fontSize:15}}>
                🚪 Sign Out
              </div>
              {signedOut?(
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:44,height:44,borderRadius:"50%",
                    background:"rgba(79,142,247,.15)",display:"flex",
                    alignItems:"center",justifyContent:"center",fontSize:22}}>✅</div>
                  <div>
                    <div style={{fontWeight:700,color:"#4f8ef7",fontSize:15}}>
                      Signed out at {signoutTime}
                    </div>
                    <div style={{fontSize:12,color:HW.muted}}>Time out recorded</div>
                  </div>
                </div>
              ):(
                <button style={{...s.btn,background:"rgba(79,142,247,.15)",color:"#4f8ef7",
                  width:"100%",padding:14,fontSize:15,
                  border:"1px solid rgba(79,142,247,.3)",opacity:loading?0.6:1}}
                  onClick={submitSignOut} disabled={loading}>
                  🚪 Record Time Out
                </button>
              )}
            </div>
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
        </div>
      )}

      {/* ══ WEEKLY REPORT TAB ══ */}
      {traineeTab==="weekly"&&(
        <div>
          <div style={{...s.card,
            background:`linear-gradient(135deg,${HW.red}15,${HW.red}05)`,
            border:`1px solid ${HW.red}40`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                  textTransform:"uppercase",marginBottom:4}}>Current Week</div>
                <div style={{fontSize:18,fontWeight:800,color:HW.text}}>{currentWeek.label}</div>
                <div style={{fontSize:12,color:HW.muted,marginTop:2}}>Sunday — Thursday</div>
              </div>
              <div style={{background:weeklySubmitted?"rgba(52,211,153,.15)":"rgba(255,165,0,.15)",
                border:weeklySubmitted?"1px solid rgba(52,211,153,.3)":"1px solid rgba(255,165,0,.3)",
                borderRadius:12,padding:"10px 14px",textAlign:"center"}}>
                <div style={{fontSize:20}}>{weeklySubmitted?"✅":"📝"}</div>
                <div style={{fontSize:11,fontWeight:700,marginTop:4,
                  color:weeklySubmitted?"#34d399":"#FFA500"}}>
                  {weeklySubmitted?"Submitted":"Pending"}
                </div>
              </div>
            </div>
          </div>

          {weeklySubmitted&&aiResult?(
            <div>
              <div style={{...s.card,border:"1px solid rgba(52,211,153,.3)"}}>
                <div style={{display:"flex",justifyContent:"space-between",
                  alignItems:"center",marginBottom:12}}>
                  <h3 style={{margin:0,color:"#34d399",fontSize:16}}>✅ Report Submitted</h3>
                  <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,
                    fontSize:12,padding:"6px 12px"}}
                    onClick={()=>setWeeklySubmitted(false)}>✏️ Edit</button>
                </div>
                <div style={{background:HW.surface2,borderRadius:10,padding:12,
                  borderLeft:`3px solid ${HW.red}`,fontSize:13,lineHeight:1.7,color:HW.text}}>
                  {weeklyText}
                </div>
              </div>
              <div style={{...s.card,border:`1px solid ${HW.red}40`}}>
                <h3 style={{marginBottom:16,color:HW.red,fontSize:16}}>🤖 AI Analysis</h3>
                <div style={{background:HW.surface2,borderRadius:12,padding:16,
                  textAlign:"center",marginBottom:16}}>
                  <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                    textTransform:"uppercase",marginBottom:8}}>Weekly KPI Score</div>
                  <div style={{fontSize:56,fontWeight:800,color:kpiColor(aiResult.kpi_score)}}>
                    {aiResult.kpi_score}
                  </div>
                  <div style={{fontSize:11,color:HW.muted}}>out of 100</div>
                </div>
                {aiResult.pie_chart&&<PieChart data={aiResult.pie_chart}/>}
                {aiResult.summary&&(
                  <div style={{background:HW.surface2,borderRadius:10,padding:12,
                    marginTop:12,borderLeft:`3px solid ${HW.red}`}}>
                    <div style={{fontSize:10,color:HW.muted,fontWeight:700,
                      marginBottom:4,textTransform:"uppercase"}}>Week Summary</div>
                    <div style={{fontSize:13,color:HW.text}}>{aiResult.summary}</div>
                  </div>
                )}
                {aiResult.talent_notes&&(
                  <div style={{background:HW.surface2,borderRadius:10,padding:12,
                    marginTop:10,borderLeft:"3px solid #FFA500"}}>
                    <div style={{fontSize:10,color:"#FFA500",fontWeight:700,
                      marginBottom:4,textTransform:"uppercase"}}>🌟 Talent Notes</div>
                    <div style={{fontSize:13,lineHeight:1.6,color:HW.text}}>{aiResult.talent_notes}</div>
                  </div>
                )}
              </div>
            </div>
          ):(
            <div style={s.card}>
              <h3 style={{marginBottom:8,color:HW.red,fontSize:16}}>📅 This Week's Tasks</h3>
              <p style={{color:HW.muted,fontSize:13,marginBottom:12,lineHeight:1.5}}>
                Write what you did each day. AI will analyze and score your report.
              </p>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
                {["Sun","Mon","Tue","Wed","Thu"].map(day=>(
                  <span key={day} style={{background:HW.surface2,borderRadius:6,
                    padding:"4px 10px",fontSize:12,color:HW.muted,
                    border:`1px solid ${HW.border}`}}>{day}</span>
                ))}
              </div>
              <label style={s.label}>Weekly Tasks *</label>
              <textarea style={{...s.input,height:200,resize:"vertical",marginBottom:14,
                fontSize:14,lineHeight:1.6}}
                placeholder="Sunday: Attended orientation with mentor...\nMonday: Completed circuit analysis...\nTuesday: Team meeting...\nWednesday: Lab session...\nThursday: Final review..."
                value={weeklyText}
                onChange={e=>setWeeklyText(e.target.value)}/>
              <label style={s.label}>📸 Proof Photo (optional)</label>
              <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:16}}>
                <button style={{...s.btn,background:HW.surface2,color:HW.text,
                  border:`1px dashed ${HW.border}`,flex:1}}
                  onClick={()=>weeklyPhotoRef.current.click()}>
                  {weeklyPhotoFile?"📷 Change Photo":"📷 Upload Proof"}
                </button>
                <input ref={weeklyPhotoRef} type="file" accept="image/*"
                  style={{display:"none"}} onChange={handleWeeklyPhoto}/>
                {weeklyPhotoPreview&&(
                  <img src={weeklyPhotoPreview} alt="proof"
                    style={{width:70,height:70,objectFit:"cover",
                      borderRadius:10,border:`2px solid ${HW.border}`}}/>
                )}
              </div>
              <button style={{...s.btn,background:HW.red,color:HW.white,
                width:"100%",padding:16,fontSize:16,
                opacity:(loading||aiLoading)?0.6:1}}
                onClick={submitWeeklyReport} disabled={loading||aiLoading}>
                {aiLoading?"🤖 AI Analyzing…":loading?"Saving…":"📅 Submit Weekly Report"}
              </button>
              {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":
                msg.startsWith("🤖")?"#FFA500":HW.red,
                fontSize:14,marginTop:12,textAlign:"center"}}>{msg}</p>}
            </div>
          )}
        </div>
      )}

      {/* ══ GOALS TAB ══ */}
      {traineeTab==="goals"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>
            {[
              {label:"Total",value:goals.length,color:HW.red},
              {label:"In Progress",value:goals.filter(g=>g.status==="in_progress").length,color:"#4f8ef7"},
              {label:"Completed",value:goals.filter(g=>g.status==="completed").length,color:"#34d399"},
              {label:"Overdue",value:goals.filter(g=>g.due_date&&
                new Date(g.due_date)<new Date()&&g.status!=="completed").length,color:"#f87171"},
            ].map((stat,i)=>(
              <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                borderRadius:12,padding:14,textAlign:"center",
                borderTop:`3px solid ${stat.color}`}}>
                <div style={{fontSize:26,fontWeight:800,color:stat.color}}>{stat.value}</div>
                <div style={{fontSize:11,color:HW.muted,marginTop:3}}>{stat.label}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:8,marginBottom:12}}>
            <button style={{...s.btn,
              background:goalFilter==="all"?HW.red:HW.surface2,
              color:goalFilter==="all"?HW.white:HW.muted,
              padding:"8px 14px",fontSize:12,whiteSpace:"nowrap"}}
              onClick={()=>setGoalFilter("all")}>All</button>
            {KRA_CATEGORIES.map(kra=>(
              <button key={kra.id}
                style={{...s.btn,
                  background:goalFilter===kra.id?kra.color:HW.surface2,
                  color:goalFilter===kra.id?HW.white:HW.muted,
                  padding:"8px 14px",fontSize:12,whiteSpace:"nowrap",
                  border:goalFilter===kra.id?"none":`1px solid ${HW.border}`}}
                onClick={()=>setGoalFilter(goalFilter===kra.id?"all":kra.id)}>
                {kra.icon} {kra.label.split(" ")[0]}
              </button>
            ))}
          </div>
          <div style={{display:"flex",justifyContent:"space-between",
            alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:14,fontWeight:700,color:HW.text}}>
              {filteredGoals.length} goal{filteredGoals.length!==1?"s":""}
            </div>
            <button style={{...s.btn,background:HW.red,color:HW.white,padding:"10px 16px"}}
              onClick={()=>setShowAddGoal(!showAddGoal)}>
              {showAddGoal?"✕ Cancel":"+ Add Goal"}
            </button>
          </div>
          {showAddGoal&&(
            <div style={{...s.card,border:`1px solid ${HW.red}40`,marginBottom:12}}>
              <h4 style={{marginBottom:14,color:HW.red,fontSize:15}}>🎯 New Goal</h4>
              <label style={s.label}>KRA Category</label>
              <select style={{...s.input,marginBottom:12}} value={newGoal.kra}
                onChange={e=>setNewGoal({...newGoal,kra:e.target.value})}>
                {KRA_CATEGORIES.map(k=>(
                  <option key={k.id} value={k.id}>{k.icon} {k.label}</option>
                ))}
              </select>
              <label style={s.label}>Goal Title *</label>
              <input style={{...s.input,marginBottom:12}}
                placeholder="e.g. Achieve 95% attendance"
                value={newGoal.goal_title}
                onChange={e=>setNewGoal({...newGoal,goal_title:e.target.value})}/>
              <label style={s.label}>Description</label>
              <textarea style={{...s.input,height:70,resize:"vertical",marginBottom:12}}
                placeholder="How do you plan to achieve this?"
                value={newGoal.description}
                onChange={e=>setNewGoal({...newGoal,description:e.target.value})}/>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                <div>
                  <label style={s.label}>Target</label>
                  <input style={s.input} type="number" value={newGoal.target_value}
                    onChange={e=>setNewGoal({...newGoal,target_value:parseFloat(e.target.value)})}/>
                </div>
                <div>
                  <label style={s.label}>Unit</label>
                  <select style={s.input} value={newGoal.unit}
                    onChange={e=>setNewGoal({...newGoal,unit:e.target.value})}>
                    <option value="%">%</option>
                    <option value="days">Days</option>
                    <option value="sessions">Sessions</option>
                    <option value="tasks">Tasks</option>
                    <option value="hours">Hours</option>
                    <option value="score">Score</option>
                  </select>
                </div>
              </div>
              <label style={s.label}>Due Date</label>
              <input style={{...s.input,marginBottom:14}} type="date" value={newGoal.due_date}
                onChange={e=>setNewGoal({...newGoal,due_date:e.target.value})}/>
              <button style={{...s.btn,background:HW.red,color:HW.white,
                width:"100%",padding:14,opacity:loading?0.6:1}}
                onClick={addGoal} disabled={loading}>
                {loading?"Saving…":"🎯 Set Goal"}
              </button>
              {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
            </div>
          )}
          {filteredGoals.length===0?(
            <div style={{...s.card,textAlign:"center",padding:32}}>
              <div style={{fontSize:36,marginBottom:10}}>🎯</div>
              <p style={{color:HW.muted,fontSize:14}}>No goals yet. Tap "+ Add Goal"!</p>
            </div>
          ):(
            filteredGoals.map(goal=>(
              <GoalCard key={goal.id} goal={goal}
                onUpdate={updateGoal} onDelete={deleteGoal} isTrainee={true}/>
            ))
          )}
          {msg&&!showAddGoal&&(
            <p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:14,marginTop:8,textAlign:"center"}}>{msg}</p>
          )}
        </div>
      )}
    </div>
  );

  // ══ MANAGEMENT DASHBOARD ══
  if(view==="mgmt") {
    const analytics=getAnalytics();
    return (
      <div style={s.page}>
        <MgmtNav/>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
          marginBottom:12,paddingTop:4}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <HuaweiLogo size={28}/>
            <div>
              <div style={{fontWeight:700,fontSize:14,color:HW.text}}>Management</div>
              <div style={{fontSize:10,color:HW.muted}}>{user?.email}</div>
            </div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:18,fontWeight:800,color:clockColor,fontFamily:"monospace"}}>
              {hh}:{mm}:{ss}
            </div>
            <div style={{fontSize:10,color:clockColor,fontWeight:700}}>
              {isLate?"⚠️ Past Deadline":"✅ Open"}
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:6,marginBottom:12}}>
          <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,
            fontSize:12,padding:"8px 12px"}}
            onClick={exportExcel}>📊 Export</button>
          <button style={{...s.btn,background:`${HW.red}20`,color:HW.red,
            fontSize:12,padding:"8px 12px",marginLeft:"auto"}}
            onClick={logout}>Sign out</button>
        </div>
        <div style={{height:1,background:HW.border,marginBottom:16}}/>

        {/* TRAINEES */}
        {mgmtTab==="trainees"&&!selected&&(
          <div>
            <h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>👥 All Trainees</h3>
            {trainees.length===0?(
              <div style={{...s.card,textAlign:"center",padding:32}}>
                <div style={{fontSize:36,marginBottom:10}}>👥</div>
                <p style={{color:HW.muted}}>No trainees yet.</p>
              </div>
            ):(
              trainees.map(t=>(
                <div key={t.id} style={{...s.card,cursor:"pointer",
                  opacity:t.status==="dropped"?0.5:1,
                  display:"flex",alignItems:"center",gap:12,padding:14}}
                  onClick={()=>openProfile(t)}>
                  <div style={{width:44,height:44,borderRadius:12,
                    background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,
                    display:"flex",alignItems:"center",justifyContent:"center",
                    fontSize:16,fontWeight:800,color:HW.white,flexShrink:0}}>
                    {t.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}
                  </div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:15,color:HW.text,
                      overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {t.full_name}
                    </div>
                    <div style={{fontSize:12,color:HW.muted,marginTop:2}}>
                      {t.department} · GPA {t.gpa||"—"}
                    </div>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",
                    alignItems:"flex-end",gap:6,flexShrink:0}}>
                    <span style={{padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,
                      background:statusColors[t.status]?.bg,color:statusColors[t.status]?.color}}>
                      {t.status}
                    </span>
                    <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,
                      fontSize:11,padding:"4px 8px"}}
                      onClick={e=>{e.stopPropagation();exportPDF(t);}}>📄 PDF</button>
                  </div>
                </div>
              ))
            )}
          </div>
        )}

        {/* TRAINEE PROFILE */}
        {mgmtTab==="trainees"&&selected&&(
          <div>
            <button style={{...s.btn,background:HW.surface2,color:HW.text,
              marginBottom:12,border:`1px solid ${HW.border}`,fontSize:13}}
              onClick={()=>{setSelected(null);setMsg("");}}>← Back</button>
            <div style={s.card}>
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:12}}>
                <div style={{width:52,height:52,borderRadius:14,
                  background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:20,fontWeight:800,color:HW.white,flexShrink:0}}>
                  {selected.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}
                </div>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:17,color:HW.text}}>{selected.full_name}</div>
                  <div style={{fontSize:12,color:HW.muted}}>
                    {selected.department} · {selected.assigned_mentor}
                  </div>
                  <span style={{padding:"3px 8px",borderRadius:20,fontSize:11,fontWeight:700,
                    marginTop:4,display:"inline-block",
                    background:statusColors[selected.status]?.bg,
                    color:statusColors[selected.status]?.color}}>
                    {selected.status}
                  </span>
                </div>
              </div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,
                  fontSize:12,padding:"8px 12px"}}
                  onClick={()=>exportPDF()}>📄 PDF</button>
                {selected.status!=="active"&&(
                  <button style={{...s.btn,background:`${HW.red}15`,color:HW.red,
                    fontSize:12,padding:"8px 12px"}}
                    onClick={()=>changeStatus("active")}>✅ Reactivate</button>
                )}
                {selected.status!=="transferred"&&(
                  <button style={{...s.btn,background:"rgba(255,165,0,.15)",color:"#FFA500",
                    fontSize:12,padding:"8px 12px"}}
                    onClick={()=>changeStatus("transferred")}>🔄 Transfer</button>
                )}
                {selected.status!=="inactive"&&(
                  <button style={{...s.btn,background:"rgba(136,136,136,.15)",color:HW.muted,
                    fontSize:12,padding:"8px 12px"}}
                    onClick={()=>changeStatus("inactive")}>⏸ Inactive</button>
                )}
                {selected.status!=="dropped"&&(
                  <button style={{...s.btn,background:"rgba(100,100,100,.2)",color:"#666",
                    fontSize:12,padding:"8px 12px"}}
                    onClick={()=>changeStatus("dropped")}>🔴 Drop</button>
                )}
              </div>
            </div>
            <div style={{display:"flex",gap:0,background:HW.surface2,
              borderRadius:12,padding:4,marginBottom:14,overflowX:"auto"}}>
              {["timeline","edit","reports","penalties","goals"].map(tab=>(
                <button key={tab} onClick={()=>setProfileTab(tab)}
                  style={{...s.btn,flex:1,padding:"10px 6px",
                    background:profileTab===tab?HW.surface:"none",
                    color:profileTab===tab?HW.text:HW.muted,
                    fontSize:10,borderRadius:8,whiteSpace:"nowrap",
                    borderBottom:profileTab===tab?`2px solid ${HW.red}`:"none"}}>
                  {tab==="timeline"?"📅 Timeline":tab==="edit"?"✏️ Edit":
                   tab==="reports"?"📋 Reports":tab==="penalties"?"⚠️ Penalties":"🎯 Goals"}
                </button>
              ))}
            </div>

            {profileTab==="timeline"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:16,fontSize:16,color:HW.text}}>Activity Timeline</h3>
                {logs.length===0?<p style={{color:HW.muted}}>No activity yet.</p>
                  :logs.map((log,i)=>(
                    <div key={log.id} style={{display:"flex",gap:12,
                      marginBottom:16,position:"relative"}}>
                      {i<logs.length-1&&(
                        <div style={{position:"absolute",left:15,top:32,width:2,
                          height:"calc(100% + 4px)",background:HW.border}}/>
                      )}
                      <div style={{width:32,height:32,borderRadius:"50%",
                        background:HW.surface2,border:`2px solid ${HW.border}`,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        fontSize:14,flexShrink:0,zIndex:1}}>
                        {eventIcons[log.event_type]||"📌"}
                      </div>
                      <div style={{flex:1,paddingTop:4}}>
                        <div style={{fontWeight:600,fontSize:13,color:HW.text}}>{log.description}</div>
                        <div style={{fontSize:11,color:HW.muted,marginTop:3}}>
                          {new Date(log.created_at).toLocaleDateString("en-GB",{
                            day:"numeric",month:"short",year:"numeric",
                            hour:"2-digit",minute:"2-digit"})}
                        </div>
                      </div>
                    </div>
                  ))
                }
              </div>
            )}

            {profileTab==="edit"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>Edit Profile</h3>
                <label style={s.label}>Department</label>
                <input style={{...s.input,marginBottom:12}} value={selected.department||""}
                  onChange={e=>setSelected({...selected,department:e.target.value,
                    _original:selected._original||{...selected}})}/>
                <label style={s.label}>Assigned Mentor</label>
                <input style={{...s.input,marginBottom:12}} value={selected.assigned_mentor||""}
                  onChange={e=>setSelected({...selected,assigned_mentor:e.target.value,
                    _original:selected._original||{...selected}})}/>
                <label style={s.label}>GPA</label>
                <input style={{...s.input,marginBottom:12}} type="number"
                  step="0.01" min="0" max="4" value={selected.gpa||""}
                  onChange={e=>setSelected({...selected,gpa:e.target.value})}/>
                <label style={s.label}>Joining Date</label>
                <input style={{...s.input,marginBottom:14}} type="date"
                  value={selected.joining_date||""}
                  onChange={e=>setSelected({...selected,joining_date:e.target.value})}/>
                <button style={{...s.btn,background:HW.red,color:HW.white,
                  width:"100%",padding:14}} onClick={saveProfile}>Save Changes</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                  fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
              </div>
            )}

            {profileTab==="reports"&&(
              <div>
                {reports.length===0?<p style={{color:HW.muted,padding:16}}>No reports yet.</p>
                  :reports.map(r=>{
                    let pie=null;
                    try{pie=r.pie_chart_json?JSON.parse(r.pie_chart_json):null;}catch(e){}
                    const isWeekly=!!r.weekly_tasks;
                    return (
                      <div key={r.id} style={{...s.card,
                        borderLeft:`4px solid ${isWeekly?"#7c5cfc":HW.red}`}}>
                        <div style={{display:"flex",justifyContent:"space-between",
                          alignItems:"center",marginBottom:10}}>
                          <div>
                            <span style={{padding:"2px 8px",borderRadius:8,fontSize:10,
                              fontWeight:700,marginRight:6,
                              background:isWeekly?"rgba(124,92,252,.15)":"rgba(207,10,44,.15)",
                              color:isWeekly?"#7c5cfc":HW.red}}>
                              {isWeekly?"📅 Weekly":"📋 Daily"}
                            </span>
                            <span style={{fontSize:13,fontWeight:600,color:HW.text}}>
                              {isWeekly?`${r.week_start} → ${r.week_end||""}`:r.report_date}
                            </span>
                          </div>
                          {r.kpi_score&&(
                            <div style={{fontSize:20,fontWeight:800,color:kpiColor(r.kpi_score)}}>
                              {r.kpi_score}<span style={{fontSize:10,color:HW.muted,fontWeight:400}}> KPI</span>
                            </div>
                          )}
                        </div>
                        {!isWeekly&&(
                          <div style={{display:"flex",gap:12,marginBottom:10,
                            background:HW.surface,borderRadius:8,padding:"8px 12px"}}>
                            <div style={{textAlign:"center",flex:1}}>
                              <div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>TIME IN</div>
                              <div style={{fontSize:14,fontWeight:800,color:"#34d399",fontFamily:"monospace"}}>{r.signin_time||"—"}</div>
                            </div>
                            <div style={{width:1,background:HW.border}}/>
                            <div style={{textAlign:"center",flex:1}}>
                              <div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>TIME OUT</div>
                              <div style={{fontSize:14,fontWeight:800,color:"#4f8ef7",fontFamily:"monospace"}}>{r.signout_time||"—"}</div>
                            </div>
                            <div style={{width:1,background:HW.border}}/>
                            <div style={{textAlign:"center",flex:1}}>
                              <div style={{fontSize:10,color:HW.muted,fontWeight:700,marginBottom:2}}>STATUS</div>
                              <div style={{fontSize:12,fontWeight:700,color:r.attended?"#34d399":HW.red}}>
                                {r.attended?"● Present":"○ Absent"}
                              </div>
                            </div>
                          </div>
                        )}
                        {isWeekly&&r.weekly_tasks&&(
                          <div style={{fontSize:13,marginBottom:10,
                            borderLeft:`3px solid #7c5cfc`,paddingLeft:10,
                            lineHeight:1.6,color:HW.muted}}>
                            {r.weekly_tasks.substring(0,200)}{r.weekly_tasks.length>200?"…":""}
                          </div>
                        )}
                        {r.penalty_applied&&(
                          <div style={{fontSize:12,color:HW.red,fontWeight:700,marginBottom:6}}>
                            ⚠️ Penalty: -{r.penalty_amount}%
                          </div>
                        )}
                        {pie&&<PieChart data={pie}/>}
                        {r.talent_notes&&(
                          <div style={{background:`${HW.red}10`,borderRadius:8,
                            padding:10,marginTop:10,borderLeft:`3px solid ${HW.red}`}}>
                            <div style={{fontSize:10,color:HW.red,fontWeight:700,marginBottom:4}}>
                              🌟 AI TALENT NOTES
                            </div>
                            <div style={{fontSize:12,lineHeight:1.6,color:HW.text}}>{r.talent_notes}</div>
                          </div>
                        )}
                      </div>
                    );
                  })
                }
              </div>
            )}

            {profileTab==="penalties"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>⚠️ Penalties</h3>
                {penalties.length===0?<p style={{color:HW.muted}}>No penalties recorded.</p>:(
                  <>
                    <div style={{background:`${HW.red}15`,border:`1px solid ${HW.red}30`,
                      borderRadius:10,padding:12,marginBottom:12}}>
                      <div style={{fontSize:13,color:HW.red,fontWeight:700}}>
                        Total: {penalties.length} × {PENALTY_PCT}% = {(penalties.length*PENALTY_PCT).toFixed(2)}% deduction
                      </div>
                    </div>
                    {penalties.map(p=>(
                      <div key={p.id} style={{display:"flex",justifyContent:"space-between",
                        alignItems:"center",background:HW.surface2,borderRadius:10,
                        borderLeft:`3px solid ${HW.red}`,padding:12,marginBottom:8}}>
                        <div>
                          <div style={{fontSize:13,fontWeight:600,color:HW.text}}>📅 {p.report_date}</div>
                          <div style={{fontSize:12,color:HW.muted,marginTop:2}}>{p.reason}</div>
                        </div>
                        <div style={{fontSize:16,fontWeight:800,color:HW.red}}>-{p.amount}%</div>
                      </div>
                    ))}
                  </>
                )}
              </div>
            )}

            {profileTab==="goals"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:12,fontSize:16,color:HW.text}}>🎯 Goals</h3>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",
                  gap:8,marginBottom:14}}>
                  {KRA_CATEGORIES.map(kra=>{
                    const kraGoals=selectedGoals.filter(g=>g.kra===kra.id);
                    const completed=kraGoals.filter(g=>g.status==="completed").length;
                    return (
                      <div key={kra.id} style={{background:HW.surface2,borderRadius:10,
                        padding:10,borderTop:`3px solid ${kra.color}`,textAlign:"center"}}>
                        <div style={{fontSize:16,marginBottom:2}}>{kra.icon}</div>
                        <div style={{fontSize:9,color:kra.color,fontWeight:700,
                          textTransform:"uppercase",marginBottom:2}}>
                          {kra.label.split(" ")[0]}
                        </div>
                        <div style={{fontSize:13,fontWeight:700,color:HW.text}}>{kraGoals.length}</div>
                        <div style={{fontSize:10,color:"#34d399"}}>{completed} done</div>
                      </div>
                    );
                  })}
                </div>
                {selectedGoals.length===0
                  ?<p style={{color:HW.muted}}>No goals set yet.</p>
                  :selectedGoals.map(goal=>(
                    <GoalCard key={goal.id} goal={goal} isTrainee={false}/>
                  ))
                }
              </div>
            )}
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
          </div>
        )}

        {/* LIVE */}
        {mgmtTab==="live"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>
              {[
                {label:"Signed In",value:liveSignins.filter(s=>s.attended).length,icon:"✅",color:"#34d399"},
                {label:"Absent",value:liveSignins.filter(s=>!s.attended).length,icon:"❌",color:HW.red},
                {label:"On Time",value:liveSignins.filter(s=>s.signin_time&&s.signin_time<=MAX_SIGNIN).length,icon:"⏰",color:"#4f8ef7"},
                {label:"Late",value:liveSignins.filter(s=>s.signin_time&&s.signin_time>MAX_SIGNIN).length,icon:"⚠️",color:"#FFA500"},
              ].map((stat,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:12,padding:14,borderTop:`3px solid ${stat.color}`}}>
                  <div style={{fontSize:20,marginBottom:6}}>{stat.icon}</div>
                  <div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div>
                  <div style={{fontSize:11,color:HW.muted,marginTop:2}}>{stat.label}</div>
                </div>
              ))}
            </div>
            <div style={s.card}>
              <div style={{display:"flex",justifyContent:"space-between",
                alignItems:"center",marginBottom:14}}>
                <h3 style={{margin:0,fontSize:16,color:HW.text}}>📡 Live Feed</h3>
                <div style={{display:"flex",alignItems:"center",gap:6,
                  background:"rgba(52,211,153,.1)",border:"1px solid rgba(52,211,153,.3)",
                  borderRadius:20,padding:"4px 10px"}}>
                  <div style={{width:7,height:7,borderRadius:"50%",
                    background:"#34d399",animation:"pulse 1.5s infinite"}}/>
                  <span style={{fontSize:11,color:"#34d399",fontWeight:700}}>LIVE</span>
                </div>
              </div>
              <style>{`
                @keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(1.3)}}
                @keyframes slideIn{from{transform:translateX(-20px);opacity:0}to{transform:translateX(0);opacity:1}}
              `}</style>
              {liveSignins.length===0?(
                <div style={{textAlign:"center",padding:32}}>
                  <div style={{fontSize:36,marginBottom:10}}>📡</div>
                  <p style={{color:HW.muted,fontSize:14}}>Waiting for sign-ins…</p>
                </div>
              ):(
                liveSignins.map((signin,i)=>(
                  <div key={signin.id} style={{display:"flex",alignItems:"center",
                    gap:12,padding:"12px 0",borderBottom:`1px solid ${HW.border}`,
                    animation:i===0?"slideIn .4s ease":"none"}}>
                    <div style={{width:40,height:40,borderRadius:"50%",
                      background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:14,fontWeight:800,color:HW.white,flexShrink:0}}>
                      {signin.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:700,fontSize:14,color:HW.text,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {signin.full_name}
                      </div>
                      <div style={{fontSize:11,color:HW.muted,marginTop:1}}>{signin.department}</div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:16,fontWeight:800,fontFamily:"monospace",
                        color:signin.signin_time>MAX_SIGNIN?HW.red:"#34d399"}}>
                        {signin.signin_time||"—"}
                      </div>
                      <span style={{fontSize:10,fontWeight:700,
                        color:signin.attended?"#34d399":HW.red}}>
                        {signin.attended?"● Present":"○ Absent"}
                      </span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )}

        {/* ANALYTICS */}
        {mgmtTab==="analytics"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:16}}>
              {[
                {label:"Active Trainees",value:analytics.active.length,icon:"👥",color:HW.red},
                {label:"Avg KPI",value:analytics.avgKpi,icon:"📊",color:"#FFA500"},
                {label:"Attendance Rate",value:`${analytics.attendanceRate}%`,icon:"✅",color:"#34d399"},
                {label:"Penalties",value:analytics.totalPenalties,icon:"⚠️",color:"#f87171"},
              ].map((stat,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:12,padding:14,borderTop:`3px solid ${stat.color}`}}>
                  <div style={{fontSize:20,marginBottom:6}}>{stat.icon}</div>
                  <div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div>
                  <div style={{fontSize:11,color:HW.muted,marginTop:2}}>{stat.label}</div>
                </div>
              ))}
            </div>
            <div style={s.card}>
              <h3 style={{marginBottom:14,fontSize:16,color:HW.text}}>🏆 KPI Leaderboard</h3>
              {analytics.traineeKpi.length===0?<p style={{color:HW.muted}}>No data yet.</p>
                :analytics.traineeKpi.slice(0,8).map((t,i)=>(
                  <div key={t.id} style={{display:"flex",alignItems:"center",gap:10,
                    padding:"10px 0",borderBottom:`1px solid ${HW.border}`}}>
                    <div style={{width:28,height:28,borderRadius:"50%",
                      background:i===0?"#FFD700":i===1?"#C0C0C0":i===2?"#CD7F32":HW.surface2,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:12,fontWeight:800,color:i<3?"#000":HW.muted,flexShrink:0}}>
                      {i+1}
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontWeight:600,fontSize:14,color:HW.text,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {t.full_name}
                      </div>
                      <div style={{fontSize:11,color:HW.muted}}>{t.department}</div>
                    </div>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:18,fontWeight:800,color:kpiColor(parseFloat(t.avgKpi))}}>
                        {t.avgKpi}
                      </div>
                      <div style={{fontSize:10,color:HW.muted}}>KPI</div>
                    </div>
                  </div>
                ))
              }
            </div>
            {analytics.atRisk.length>0&&(
              <div style={{...s.card,border:"1px solid rgba(248,113,113,.3)"}}>
                <h3 style={{marginBottom:12,color:"#f87171",fontSize:16}}>⚠️ At Risk</h3>
                {analytics.atRisk.map(t=>(
                  <div key={t.id} style={{background:HW.surface2,borderRadius:10,
                    padding:12,marginBottom:8,borderLeft:"3px solid #f87171"}}>
                    <div style={{fontWeight:700,fontSize:14,color:HW.text}}>{t.full_name}</div>
                    <div style={{fontSize:12,color:HW.muted,marginBottom:6}}>{t.department}</div>
                    <div style={{display:"flex",gap:16}}>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontSize:18,fontWeight:800,color:"#f87171"}}>{t.avgKpi}</div>
                        <div style={{fontSize:10,color:HW.muted}}>Avg KPI</div>
                      </div>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontSize:18,fontWeight:800,color:"#f87171"}}>{t.penalties}</div>
                        <div style={{fontSize:10,color:HW.muted}}>Penalties</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* OKR */}
        {mgmtTab==="okr"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",marginBottom:14}}>
              <h3 style={{margin:0,fontSize:16,color:HW.text}}>🎯 OKR Tracking</h3>
              <button style={{...s.btn,background:HW.red,color:HW.white,
                padding:"8px 14px",fontSize:13}}
                onClick={()=>setShowAddOkr(!showAddOkr)}>
                {showAddOkr?"Cancel":"+ Add"}
              </button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:14}}>
              {[
                {label:"Total",value:okrs.length,color:HW.red},
                {label:"On Track",value:okrs.filter(o=>(o.current/o.target)>=.8).length,color:"#34d399"},
                {label:"In Progress",value:okrs.filter(o=>(o.current/o.target)>=.5&&(o.current/o.target)<.8).length,color:"#FFA500"},
                {label:"Behind",value:okrs.filter(o=>(o.current/o.target)<.5).length,color:"#f87171"},
              ].map((s2,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:12,padding:14,textAlign:"center",borderTop:`3px solid ${s2.color}`}}>
                  <div style={{fontSize:24,fontWeight:800,color:s2.color}}>{s2.value}</div>
                  <div style={{fontSize:11,color:HW.muted,marginTop:2}}>{s2.label}</div>
                </div>
              ))}
            </div>
            {showAddOkr&&(
              <div style={{...s.card,border:`1px solid ${HW.red}40`,marginBottom:14}}>
                <h4 style={{marginBottom:14,color:HW.red,fontSize:15}}>Add OKR</h4>
                <label style={s.label}>Department</label>
                <input style={{...s.input,marginBottom:12}} placeholder="e.g. Engineering"
                  value={newOkr.department}
                  onChange={e=>setNewOkr({...newOkr,department:e.target.value})}/>
                <label style={s.label}>Objective</label>
                <input style={{...s.input,marginBottom:12}} placeholder="e.g. Improve Skills"
                  value={newOkr.objective}
                  onChange={e=>setNewOkr({...newOkr,objective:e.target.value})}/>
                <label style={s.label}>Key Result</label>
                <input style={{...s.input,marginBottom:12}} placeholder="e.g. Complete 20 sessions"
                  value={newOkr.key_result}
                  onChange={e=>setNewOkr({...newOkr,key_result:e.target.value})}/>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                  <div>
                    <label style={s.label}>Target</label>
                    <input style={s.input} type="number" value={newOkr.target}
                      onChange={e=>setNewOkr({...newOkr,target:parseFloat(e.target.value)})}/>
                  </div>
                  <div>
                    <label style={s.label}>Unit</label>
                    <select style={s.input} value={newOkr.unit}
                      onChange={e=>setNewOkr({...newOkr,unit:e.target.value})}>
                      <option value="%">%</option>
                      <option value="sessions">Sessions</option>
                      <option value="reports">Reports</option>
                      <option value="tasks">Tasks</option>
                    </select>
                  </div>
                </div>
                <label style={s.label}>Due Date</label>
                <input style={{...s.input,marginBottom:14}} type="date" value={newOkr.due_date}
                  onChange={e=>setNewOkr({...newOkr,due_date:e.target.value})}/>
                <button style={{...s.btn,background:HW.red,color:HW.white,
                  width:"100%",padding:14}} onClick={addOkr}>Add OKR</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                  fontSize:13,marginTop:8,textAlign:"center"}}>{msg}</p>}
              </div>
            )}
            {[...new Set(okrs.map(o=>o.department))].map(dept=>(
              <div key={dept} style={s.card}>
                <h4 style={{marginBottom:12,color:HW.red,fontSize:15}}>🏢 {dept}</h4>
                {okrs.filter(o=>o.department===dept).map(okr=>(
                  <OKRBar key={okr.id} okr={okr} onUpdate={updateOkr}/>
                ))}
              </div>
            ))}
            {okrs.length===0&&(
              <div style={{...s.card,textAlign:"center",padding:32}}>
                <div style={{fontSize:36,marginBottom:10}}>🎯</div>
                <p style={{color:HW.muted}}>No OKRs yet. Tap "+ Add"!</p>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }
}