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

const HW = {
  red:"#CF0A2C", darkRed:"#A00820", black:"#1A1A1A",
  dark:"#0D0D0D", surface:"#1E1E1E", surface2:"#2A2A2A",
  border:"#333333", text:"#F5F5F5", muted:"#888888", white:"#FFFFFF",
};

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
        `Analyze this trainee daily report and respond ONLY with JSON:
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
    <div style={{display:"flex",alignItems:"center",gap:24,flexWrap:"wrap"}}>
      <svg viewBox="-1.1 -1.1 2.2 2.2" width="160" height="160">
        {slices.map((s,i)=>(
          <path key={i} d={s.d} fill={s.color} stroke={HW.surface} strokeWidth="0.03"/>
        ))}
        <circle cx="0" cy="0" r="0.55" fill={HW.surface}/>
      </svg>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {slices.map((s,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8,fontSize:13}}>
            <div style={{width:10,height:10,borderRadius:"50%",background:s.color,flexShrink:0}}/>
            <span style={{color:HW.muted}}>{s.label}</span>
            <span style={{fontWeight:700,color:s.color,marginLeft:"auto",paddingLeft:8}}>{s.pct}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function BarChart({data}) {
  const max=Math.max(...data.map(d=>d.value),1);
  return (
    <div style={{display:"flex",alignItems:"flex-end",gap:8,height:120,padding:"0 4px"}}>
      {data.map((d,i)=>(
        <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
          <div style={{fontSize:10,color:HW.muted,fontWeight:700}}>{d.value}</div>
          <div style={{width:"100%",background:HW.red,borderRadius:"4px 4px 0 0",
            height:`${(d.value/max)*80}px`,minHeight:4}}/>
          <div style={{fontSize:9,color:HW.muted,textAlign:"center",
            whiteSpace:"nowrap",overflow:"hidden",maxWidth:40}}>{d.label}</div>
        </div>
      ))}
    </div>
  );
}

function OKRBar({okr,onUpdate}) {
  const pct=Math.min((okr.current/okr.target)*100,100).toFixed(0);
  const color=pct>=80?HW.red:pct>=50?"#FFA500":"#666";
  return (
    <div style={{background:HW.surface2,borderRadius:12,padding:16,marginBottom:12}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
        <div style={{flex:1}}>
          <div style={{fontSize:12,color:HW.muted,marginBottom:2}}>{okr.department}</div>
          <div style={{fontWeight:600,fontSize:14,marginBottom:2}}>{okr.objective}</div>
          <div style={{fontSize:13,color:HW.muted}}>{okr.key_result}</div>
        </div>
        <div style={{textAlign:"right",marginLeft:12}}>
          <div style={{fontSize:22,fontWeight:800,color}}>{pct}%</div>
          <div style={{fontSize:11,color:HW.muted}}>{okr.current}/{okr.target} {okr.unit}</div>
        </div>
      </div>
      <div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden"}}>
        <div style={{height:"100%",borderRadius:10,background:color,
          width:`${pct}%`,transition:"width .6s ease"}}/>
      </div>
      {okr.due_date&&(
        <div style={{fontSize:11,color:HW.muted,marginTop:6}}>
          Due: {new Date(okr.due_date).toLocaleDateString("en-GB",
            {day:"numeric",month:"short",year:"numeric"})}
        </div>
      )}
      {onUpdate&&(
        <div style={{display:"flex",gap:8,marginTop:10,alignItems:"center"}}>
          <input type="number" defaultValue={okr.current}
            style={{background:HW.surface,border:`1px solid ${HW.border}`,
              color:HW.text,borderRadius:6,padding:"4px 8px",width:80,fontSize:13}}
            onBlur={e=>onUpdate(okr.id,parseFloat(e.target.value))}/>
          <span style={{fontSize:12,color:HW.muted}}>Update progress</span>
        </div>
      )}
    </div>
  );
}

function GoalCard({goal,onUpdate,onDelete,isTrainee}) {
  const kra=KRA_CATEGORIES.find(k=>k.id===goal.kra)||KRA_CATEGORIES[0];
  const pct=Math.min((goal.current_value/goal.target_value)*100,100).toFixed(0);
  const isOverdue=goal.due_date&&new Date(goal.due_date)<new Date()&&goal.status!=="completed";
  const statusColors={
    not_started:{bg:"rgba(136,136,136,.15)",color:"#888"},
    in_progress:{bg:"rgba(79,142,247,.15)",color:"#4f8ef7"},
    completed:{bg:"rgba(52,211,153,.15)",color:"#34d399"},
    overdue:{bg:"rgba(248,113,113,.15)",color:"#f87171"},
  };
  const status=isOverdue&&goal.status!=="completed"?"overdue":goal.status;
  return (
    <div style={{background:HW.surface2,borderRadius:14,padding:18,
      marginBottom:14,borderLeft:`4px solid ${kra.color}`}}>
      <div style={{display:"flex",justifyContent:"space-between",
        alignItems:"flex-start",marginBottom:12}}>
        <div style={{flex:1}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
            <span style={{fontSize:18}}>{kra.icon}</span>
            <span style={{fontSize:11,color:kra.color,fontWeight:700,
              textTransform:"uppercase"}}>{kra.label}</span>
          </div>
          <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>{goal.goal_title}</div>
          {goal.description&&<div style={{fontSize:13,color:HW.muted}}>{goal.description}</div>}
        </div>
        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
          <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,
            background:statusColors[status]?.bg,color:statusColors[status]?.color}}>
            {status==="not_started"?"⬜ Not Started":status==="in_progress"?"🔵 In Progress":
             status==="completed"?"✅ Completed":"🔴 Overdue"}
          </span>
          {goal.due_date&&(
            <div style={{fontSize:11,color:isOverdue?HW.red:HW.muted}}>
              Due: {new Date(goal.due_date).toLocaleDateString("en-GB",
                {day:"numeric",month:"short",year:"numeric"})}
            </div>
          )}
        </div>
      </div>
      <div style={{marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}>
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
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <input type="number" placeholder="Update value" id={`prog_${goal.id}`}
            style={{background:HW.surface,border:`1px solid ${HW.border}`,
              color:HW.text,borderRadius:6,padding:"6px 10px",width:130,fontSize:13}}/>
          <button style={{padding:"6px 14px",borderRadius:6,border:"none",
            background:kra.color,color:HW.white,fontWeight:700,fontSize:12,cursor:"pointer"}}
            onClick={()=>{
              const input=document.getElementById(`prog_${goal.id}`);
              if(input&&input.value) onUpdate(goal.id,parseFloat(input.value));
            }}>Update</button>
          {parseFloat(pct)>=100&&(
            <button style={{padding:"6px 14px",borderRadius:6,border:"none",
              background:"rgba(52,211,153,.15)",color:"#34d399",
              fontWeight:700,fontSize:12,cursor:"pointer"}}
              onClick={()=>onUpdate(goal.id,goal.target_value,"completed")}>
              ✅ Mark Complete
            </button>
          )}
          {onDelete&&(
            <button style={{padding:"6px 14px",borderRadius:6,border:"none",
              background:"rgba(248,113,113,.1)",color:"#f87171",
              fontWeight:700,fontSize:12,cursor:"pointer"}}
              onClick={()=>onDelete(goal.id)}>🗑 Delete</button>
          )}
        </div>
      )}
    </div>
  );
}

function ReminderPopup({onDismiss}) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",
      display:"flex",alignItems:"center",justifyContent:"center",
      zIndex:999,backdropFilter:"blur(4px)"}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,
        borderRadius:20,padding:40,maxWidth:380,textAlign:"center",
        boxShadow:`0 0 60px rgba(207,10,44,.3)`}}>
        <div style={{fontSize:52,marginBottom:12}}>⏰</div>
        <h3 style={{color:HW.red,fontSize:22,marginBottom:8}}>Task Submission Reminder</h3>
        <p style={{color:HW.muted,fontSize:14,lineHeight:1.6,marginBottom:20}}>
          It is <strong style={{color:HW.text}}>4:30 PM</strong> — please submit your daily report!
        </p>
        <button onClick={onDismiss} style={{background:HW.red,color:HW.white,border:"none",
          borderRadius:8,padding:"12px 32px",fontWeight:700,fontSize:14,
          cursor:"pointer",width:"100%"}}>
          Got it — I'll submit now!
        </button>
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
        `${greeting}, ${name}! Welcome to TraineeOS. ${quote.english}. Have a nice day!`
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
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.85)",
      display:"flex",alignItems:"center",justifyContent:"center",
      zIndex:1000,backdropFilter:"blur(8px)"}}>
      <div style={{background:HW.surface,border:`2px solid ${HW.red}`,
        borderRadius:24,padding:44,maxWidth:480,width:"90%",textAlign:"center",
        boxShadow:`0 0 80px rgba(207,10,44,.25)`,
        animation:"popIn .4s cubic-bezier(.34,1.56,.64,1)"}}>
        <style>{`
          @keyframes popIn{from{transform:scale(.8);opacity:0}to{transform:scale(1);opacity:1}}
          @keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-8px)}}
        `}</style>
        <div style={{animation:"float 3s ease-in-out infinite",marginBottom:16}}>
          <HuaweiLogo size={64}/>
        </div>
        <div style={{fontSize:13,color:HW.muted,marginBottom:4,
          letterSpacing:".1em",textTransform:"uppercase"}}>{greetingAr}</div>
        <h2 style={{fontSize:32,fontWeight:800,color:HW.white,margin:"0 0 4px"}}>
          {greeting},
        </h2>
        <h2 style={{fontSize:36,fontWeight:800,color:HW.red,margin:"0 0 24px"}}>
          {name}! 👋
        </h2>
        <div style={{height:1,background:HW.border,marginBottom:24}}/>
        <div style={{background:HW.surface2,borderRadius:16,padding:20,
          marginBottom:20,border:`1px solid ${HW.border}`}}>
          <div style={{fontSize:11,color:HW.red,fontWeight:700,
            textTransform:"uppercase",letterSpacing:".1em",marginBottom:12}}>
            🈲 Today's Motivation
          </div>
          <div style={{fontSize:26,fontWeight:800,color:HW.white,marginBottom:10,
            lineHeight:1.4,fontFamily:"serif",letterSpacing:4}}>{quote.chinese}</div>
          <div style={{fontSize:14,color:HW.muted,lineHeight:1.7,
            fontStyle:"italic",marginBottom:8}}>"{quote.english}"</div>
          <div style={{fontSize:12,color:HW.red,fontWeight:600}}>— {quote.author}</div>
        </div>
        <div style={{fontSize:16,color:"#34d399",fontWeight:700,marginBottom:20}}>
          🌟 Have a nice day!
        </div>
        <div style={{fontSize:13,color:HW.muted,marginBottom:24}}>
          {now.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",
            month:"long",year:"numeric"})}
        </div>
        <button onClick={onDismiss}
          style={{background:HW.red,color:HW.white,border:"none",
            borderRadius:12,padding:"14px 40px",fontWeight:800,
            fontSize:16,cursor:"pointer",width:"100%"}}
          onMouseOver={e=>e.target.style.background=HW.darkRed}
          onMouseOut={e=>e.target.style.background=HW.red}>
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
  const [photoFile,setPhotoFile]     = useState(null);
  const [photoPreview,setPhotoPreview] = useState(null);
  const [geoStatus,setGeoStatus]     = useState("idle");
  const [geoMsg,setGeoMsg]           = useState("");
  const [locationOk,setLocationOk]   = useState(false);
  const [showReminder,setShowReminder] = useState(false);
  const [showGreeting,setShowGreeting] = useState(false);
  const [traineeName,setTraineeName] = useState("");
  const [trafficCount,setTrafficCount] = useState(0);
  const [faceCapture,setFaceCapture] = useState(null);
  const [faceCaptured,setFaceCaptured] = useState(false);
  const [facePreviewUrl,setFacePreviewUrl] = useState(null);
  const [cameraActive,setCameraActive] = useState(false);
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

  const [setupProfile,setSetupProfile] = useState({
    full_name:"",civil_id:"",phone_number:"",
    department:"",assigned_mentor:"",gpa:"",
    date_of_birth:"",gender:"",nationality:"Omani",
  });

  const videoRef=useRef(),canvasRef=useRef(),streamRef=useRef();
  const excuseRef=useRef(),photoRef=useRef();

  const [profile,setProfile] = useState({
    full_name:"",civil_id:"",phone_number:"",
    department:"",assigned_mentor:"",gpa:"",
  });
  const [attendance,setAttendance] = useState({
    report_date:new Date().toISOString().split("T")[0],
    attended:false,
  });
  const [taskReport,setTaskReport] = useState({
    report_text:"",report_date:new Date().toISOString().split("T")[0],
  });

  const s = {
    page:{minHeight:"100vh",background:HW.dark,color:HW.text,
      fontFamily:"sans-serif",padding:32},
    card:{background:HW.surface,border:`1px solid ${HW.border}`,
      borderRadius:16,padding:24,marginBottom:20},
    input:{background:HW.surface2,border:`1px solid ${HW.border}`,
      color:HW.text,borderRadius:8,padding:"9px 13px",width:"100%",
      fontFamily:"inherit",fontSize:14,boxSizing:"border-box"},
    label:{fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",
      letterSpacing:".06em",display:"block",marginBottom:5},
    btn:{padding:"10px 20px",borderRadius:8,border:"none",
      fontWeight:700,cursor:"pointer",fontSize:13},
    th:{textAlign:"left",padding:"10px 14px",fontSize:11,fontWeight:700,
      color:HW.muted,textTransform:"uppercase",borderBottom:`1px solid ${HW.border}`},
    td:{padding:"12px 14px",fontSize:14,borderBottom:`1px solid rgba(51,51,51,.6)`},
  };

  const statusColors={
    active:{bg:"rgba(207,10,44,.15)",color:HW.red},
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
          new Notification("TraineeOS",{body:"4:30 PM — Submit your daily report!"});
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
        if(!data.profile_completed){
          setView("setup");
        } else {
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
            // Check if already signed out today
            const today=new Date().toISOString().split("T")[0];
            const{data:todayReport}=await supabase.from("daily_reports")
              .select("signout_time").eq("trainee_id",data.trainee_id)
              .eq("report_date",today).single();
            if(todayReport?.signout_time){
              setSignedOut(true);setSignoutTime(todayReport.signout_time);
            }
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
    setAiResult(null);setPhotoFile(null);setPhotoPreview(null);
    setLocationOk(false);setGeoStatus("idle");setGeoMsg("");
    setFaceCapture(null);setFaceCaptured(false);setFacePreviewUrl(null);
    stopCamera();setTraineeId(null);setGoals([]);setLiveSignins([]);
    setShowGreeting(false);setTraineeName("");
    setSignedOut(false);setSignoutTime("");
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
          setGeoMsg(`❌ You are ${Math.round(dist)}m away. Must be within ${WORK_RADIUS}m.`);
        }
      },
      ()=>{setGeoStatus("error");setGeoMsg("❌ Location denied.");},
      {enableHighAccuracy:true,timeout:10000}
    );
  }

  async function startCamera(){
    try{
      const stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:"user"},audio:false});
      streamRef.current=stream;
      if(videoRef.current) videoRef.current.srcObject=stream;
      setCameraActive(true);
    }catch(e){setMsg("❌ Camera access denied.");}
  }

  function stopCamera(){
    if(streamRef.current){streamRef.current.getTracks().forEach(t=>t.stop());streamRef.current=null;}
    setCameraActive(false);
  }

  function captureface(){
    if(!videoRef.current||!canvasRef.current) return;
    const canvas=canvasRef.current,video=videoRef.current;
    canvas.width=video.videoWidth;canvas.height=video.videoHeight;
    canvas.getContext("2d").drawImage(video,0,0);
    const previewUrl=canvas.toDataURL("image/jpeg",0.8);
    setFacePreviewUrl(previewUrl);
    canvas.toBlob(blob=>{
      setFaceCapture(blob);setFaceCaptured(true);stopCamera();
      setMsg("✅ Face captured successfully!");
    },"image/jpeg",0.8);
  }

  function handleProofPhoto(e){const f=e.target.files[0];if(!f)return;setPhotoFile(f);setPhotoPreview(URL.createObjectURL(f));}
  function handleExcusePhoto(e){const f=e.target.files[0];if(!f)return;setExcusePhoto(f);setExcusePreview(URL.createObjectURL(f));}

  async function uploadFile(bucket,path,file){
    const{error}=await supabase.storage.from(bucket).upload(path,file);
    if(error) return null;
    return supabase.storage.from(bucket).getPublicUrl(path).data.publicUrl;
  }

  async function submitAttendance(){
    if(!locationOk){setMsg("📍 Please verify your location first.");return;}
    if(!faceCaptured){setMsg("🤳 Please complete Face ID first.");return;}
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
    const faceUrl=faceCapture
      ?await uploadFile("report-photos",`${tid}/face_${ts}.jpg`,faceCapture):null;
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
      face_capture_url:faceUrl,
      excuse_type:excuseType||null,excuse_text:excuseText||null,
      excuse_photo_url:excusePhotoUrl,traffic_excuse:excuseType==="traffic",
      penalty_applied:penaltyApplied,penalty_amount:penaltyApplied?PENALTY_PCT:0,
    },{onConflict:"trainee_id,report_date"});
    setMsg(penaltyApplied
      ?`✅ Attendance recorded — ⚠️ Penalty of ${PENALTY_PCT}% applied.`
      :"✅ Attendance recorded successfully!");
    setLoading(false);
  }

  // ── Sign Out ─────────────────────────────────────────
  async function submitSignOut(){
    if(!traineeId){setMsg("Please submit attendance first.");return;}
    setLoading(true);setMsg("");
    const now=new Date();
    const timeStr=`${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    const today=new Date().toISOString().split("T")[0];
    const{error}=await supabase.from("daily_reports")
      .update({signout_time:timeStr})
      .eq("trainee_id",traineeId)
      .eq("report_date",today);
    if(error){setMsg("Error: "+error.message);}
    else{
      setSignedOut(true);setSignoutTime(timeStr);
      setMsg(`✅ Sign out recorded at ${timeStr}`);
    }
    setLoading(false);
  }

  async function submitTasks(){
    if(!taskReport.report_text){setMsg("Please write your daily tasks.");return;}
    setLoading(true);setAiLoading(true);setMsg("");setAiResult(null);
    const tid=traineeId;
    if(!tid){setMsg("Please submit attendance first.");setLoading(false);setAiLoading(false);return;}
    const ts=Date.now();
    const photoUrl=photoFile
      ?await uploadFile("report-photos",`${tid}/proof_${ts}.jpg`,photoFile):null;
    setMsg("🤖 AI is analyzing your report…");
    let ai=null;
    try{ai=await analyzeReport(taskReport.report_text);setAiResult(ai);}
    catch(e){console.error(e);}
    await supabase.from("daily_reports").upsert({
      trainee_id:tid,report_date:taskReport.report_date,
      report_text:taskReport.report_text,photo_url:photoUrl,
      kpi_score:ai?.kpi_score||null,
      pie_chart_json:ai?.pie_chart?JSON.stringify(ai.pie_chart):null,
      talent_notes:ai?.talent_notes||null,
    },{onConflict:"trainee_id,report_date"});
    setMsg("✅ Tasks submitted and AI analysis complete!");
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
    setMsg("✅ Profile saved!");fetchTrainees();fetchLogs(selected.id);
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

  async function exportPDF(trainee){
    const doc=new jsPDF();const t=trainee||selected;
    doc.setFillColor(207,10,44);doc.rect(0,0,210,40,"F");
    doc.setTextColor(255,255,255);doc.setFontSize(20);doc.setFont("helvetica","bold");
    doc.text("TraineeOS — Weekly Performance Report",14,18);
    doc.setFontSize(11);doc.setFont("helvetica","normal");
    doc.text(`Generated: ${new Date().toLocaleDateString("en-GB")}`,14,30);
    doc.setTextColor(0,0,0);doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text("Trainee Information",14,52);
    autoTable(doc,{startY:56,
      head:[["Field","Details"]],
      body:[["Full Name",t.full_name||"—"],["Civil ID",t.civil_id||"—"],
        ["Department",t.department||"—"],["Mentor",t.assigned_mentor||"—"],
        ["GPA",t.gpa?.toString()||"—"],["Status",t.status||"—"],
        ["Joining Date",t.joining_date||"—"],
        ["Quitting Date",t.quitting_date||"Still Active"],
        ["Days in Program",t.joining_date
          ?Math.floor((new Date()-new Date(t.joining_date))/(1000*60*60*24))+" days":"—"]],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    const reportsY=doc.lastAutoTable.finalY+10;
    doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text("Weekly Reports",14,reportsY);
    const{data:reps}=await supabase.from("daily_reports").select("*")
      .eq("trainee_id",t.id).order("report_date",{ascending:false}).limit(7);
    if(reps?.length>0){
      autoTable(doc,{startY:reportsY+4,
        head:[["Date","Attended","Time In","Time Out","KPI","Penalty"]],
        body:reps.map(r=>[r.report_date,r.attended?"✓":"✗",
          r.signin_time||"—",r.signout_time||"—",
          r.kpi_score||"—",r.penalty_applied?`-${r.penalty_amount}%`:"None"]),
        headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
        alternateRowStyles:{fillColor:[245,245,245]},
      });
    }
    const penY=doc.lastAutoTable?.finalY+10||reportsY+60;
    doc.setFontSize(14);doc.setFont("helvetica","bold");doc.text("Penalties",14,penY);
    const{data:pens}=await supabase.from("penalties").select("*").eq("trainee_id",t.id);
    autoTable(doc,{startY:penY+4,
      head:[["Date","Reason","Deduction"]],
      body:pens?.length>0?pens.map(p=>[p.report_date,p.reason,`-${p.amount}%`])
        :[["—","No penalties","—"]],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    doc.setTextColor(207,10,44);doc.setFontSize(11);doc.setFont("helvetica","bold");
    doc.text(`Total Deduction: ${pens?(pens.length*PENALTY_PCT).toFixed(2):0}%`,
      14,doc.lastAutoTable.finalY+6);
    doc.setTextColor(150,150,150);doc.setFontSize(9);doc.setFont("helvetica","normal");
    doc.text("TraineeOS • Powered by Huawei • Confidential",14,285);
    doc.save(`${t.full_name}_weekly_report.pdf`);
  }

  async function exportExcel(){
    setMsg("📊 Preparing Excel export…");
    const{data:allT}=await supabase.from("trainees").select("*").order("full_name");
    const{data:allR}=await supabase.from("daily_reports").select("*").order("report_date");
    const{data:allP}=await supabase.from("penalties").select("*").order("created_at");
    const{data:allL}=await supabase.from("trainee_logs").select("*").order("created_at");
    const{data:allG}=await supabase.from("goals").select("*").order("created_at");
    const wb=XLSX.utils.book_new();

    // ── Trainees Sheet ───────────────────────────────
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allT||[]).map(t=>({
        "Full Name":t.full_name,"Civil ID":t.civil_id,"Phone":t.phone_number,
        "Department":t.department,"Mentor":t.assigned_mentor,"GPA":t.gpa,
        "Gender":t.gender||"—","Nationality":t.nationality||"—",
        "DOB":t.date_of_birth||"—","Status":t.status,
        "Joining Date":t.joining_date,"Quitting Date":t.quitting_date||"Still Active",
        "Days":t.joining_date
          ?Math.floor((new Date()-new Date(t.joining_date))/(1000*60*60*24)):"—",
      }))),"Trainees");

    // ── Daily Attendance Sheet (with Time In + Time Out) ──
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allR||[]).map(r=>{
        const trainee=(allT||[]).find(t=>t.id===r.trainee_id);
        const dateObj=new Date(r.report_date+"T00:00:00");
        const dayName=dateObj.toLocaleDateString("en-GB",{weekday:"long"});
        const dateFormatted=dateObj.toLocaleDateString("en-GB",{
          day:"numeric",month:"long",year:"numeric"});
        return{
          "Full Name":trainee?.full_name||"—",
          "Department":trainee?.department||"—",
          "Day":dayName,
          "Date":dateFormatted,
          "Attended":r.attended?"Yes":"No",
          "Time In":r.signin_time||"—",
          "Time Out":r.signout_time||"—",
          "Duration":r.signin_time&&r.signout_time?(()=>{
            const [ih,im]=r.signin_time.split(":").map(Number);
            const [oh,om]=r.signout_time.split(":").map(Number);
            const mins=(oh*60+om)-(ih*60+im);
            return mins>0?`${Math.floor(mins/60)}h ${mins%60}m`:"—";
          })():"—",
          "KPI Score":r.kpi_score||"—",
          "Penalty":r.penalty_applied?`-${r.penalty_amount}%`:"None",
          "Excuse":r.excuse_type||"—",
          "Report":r.report_text||"—",
        };
      })),"Daily Attendance");

    // ── Penalties Sheet ──────────────────────────────
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allP||[]).map(p=>({
        "Full Name":(allT||[]).find(t=>t.id===p.trainee_id)?.full_name||"—",
        "Date":p.report_date,"Reason":p.reason,"Deduction":`-${p.amount}%`,
      }))),"Penalties");

    // ── Goals Sheet ──────────────────────────────────
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allG||[]).map(g=>({
        "Full Name":(allT||[]).find(t=>t.id===g.trainee_id)?.full_name||"—",
        "KRA":g.kra,"Goal":g.goal_title,"Target":g.target_value,
        "Current":g.current_value,"Unit":g.unit,"Status":g.status,
        "Due Date":g.due_date||"—",
        "Progress":Math.min((g.current_value/g.target_value)*100,100).toFixed(0)+"%",
      }))),"Goals");

    // ── Summary Sheet ────────────────────────────────
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(
      (allT||[]).map(t=>{
        const tr=(allR||[]).filter(r=>r.trainee_id===t.id);
        const tp=(allP||[]).filter(p=>p.trainee_id===t.id);
        const tg=(allG||[]).filter(g=>g.trainee_id===t.id);
        const avg=tr.filter(r=>r.kpi_score).reduce((a,r,_,arr)=>a+r.kpi_score/arr.length,0);
        return{
          "Full Name":t.full_name,"Department":t.department,"Status":t.status,
          "Days Present":tr.filter(r=>r.attended).length,
          "Days Absent":tr.filter(r=>!r.attended).length,
          "Avg KPI":avg?avg.toFixed(1):"—","Penalties":tp.length,
          "Goals Set":tg.length,
          "Goals Completed":tg.filter(g=>g.status==="completed").length,
          "Total Deduction":`${(tp.length*PENALTY_PCT).toFixed(2)}%`,
        };
      })),"Summary");

    XLSX.writeFile(wb,`TraineeOS_Attendance_${new Date().toISOString().split("T")[0]}.xlsx`);
    setMsg("✅ Excel exported successfully!");
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

  // ══════════════════════════════════════════════════
  // LOGIN
  // ══════════════════════════════════════════════════
  if(view==="login") return (
    <div style={{...s.page,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{...s.card,width:420}}>
        <div style={{textAlign:"center",marginBottom:28,paddingBottom:24,
          borderBottom:`1px solid ${HW.border}`}}>
          <HuaweiLogo size={56}/>
          <h2 style={{margin:"12px 0 4px",fontSize:24}}>TraineeOS</h2>
          <p style={{color:HW.red,fontSize:12,fontWeight:700,
            letterSpacing:".1em",textTransform:"uppercase"}}>Powered by Huawei</p>
        </div>
        <label style={s.label}>Email</label>
        <input style={{...s.input,marginBottom:14}} value={email}
          onChange={e=>setEmail(e.target.value)} placeholder="you@example.com"/>
        <label style={s.label}>Password</label>
        <input style={{...s.input,marginBottom:20}} type="password"
          value={password} onChange={e=>setPassword(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&login()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,
          width:"100%",padding:13,opacity:loading?0.6:1}}
          onClick={login} disabled={loading}>
          {loading?"Signing in…":"Sign In →"}
        </button>
        <div style={{display:"flex",alignItems:"center",gap:10,margin:"20px 0"}}>
          <div style={{flex:1,height:1,background:HW.border}}/>
          <span style={{color:HW.muted,fontSize:12}}>OR</span>
          <div style={{flex:1,height:1,background:HW.border}}/>
        </div>
        <button style={{...s.btn,background:HW.surface2,color:HW.text,
          width:"100%",padding:13,border:`1px solid ${HW.border}`}}
          onClick={()=>{setView("signup");setMsg("");}}>
          Create New Trainee Account
        </button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
          fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════
  // SIGNUP
  // ══════════════════════════════════════════════════
  if(view==="signup") return (
    <div style={{...s.page,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{...s.card,width:420}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <HuaweiLogo size={40}/>
          <h2 style={{margin:"10px 0 4px"}}>Create Account</h2>
          <p style={{color:HW.muted,fontSize:13}}>Register as a new trainee</p>
        </div>
        <label style={s.label}>Email Address</label>
        <input style={{...s.input,marginBottom:14}} value={email}
          onChange={e=>setEmail(e.target.value)} placeholder="yourname@example.com"/>
        <label style={s.label}>Password</label>
        <input style={{...s.input,marginBottom:14}} type="password"
          value={password} onChange={e=>setPassword(e.target.value)}
          placeholder="Minimum 6 characters"/>
        <label style={s.label}>Confirm Password</label>
        <input style={{...s.input,marginBottom:20}} type="password"
          value={confirmPwd} onChange={e=>setConfirmPwd(e.target.value)}
          placeholder="Repeat your password"
          onKeyDown={e=>e.key==="Enter"&&signup()}/>
        <button style={{...s.btn,background:HW.red,color:HW.white,
          width:"100%",padding:13,opacity:loading?0.6:1}}
          onClick={signup} disabled={loading}>
          {loading?"Creating account…":"Create Account →"}
        </button>
        <button style={{...s.btn,background:"none",color:HW.muted,
          width:"100%",padding:13,marginTop:8}}
          onClick={()=>{setView("login");setMsg("");}}>← Back to Login</button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
          fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════
  // PROFILE SETUP
  // ══════════════════════════════════════════════════
  if(view==="setup") return (
    <div style={{...s.page,display:"flex",alignItems:"center",
      justifyContent:"center",padding:"32px 16px"}}>
      <div style={{width:"100%",maxWidth:600}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <HuaweiLogo size={56}/>
          <h2 style={{margin:"16px 0 8px",fontSize:26}}>Welcome to TraineeOS! 🎉</h2>
          <p style={{color:HW.muted,fontSize:14}}>
            Please complete your profile. You only need to do this once.
          </p>
        </div>
        <div style={{display:"flex",alignItems:"center",
          justifyContent:"center",gap:8,marginBottom:32}}>
          {["Personal Info","Academic","Contact"].map((step,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
              <div style={{width:28,height:28,borderRadius:"50%",background:HW.red,
                display:"flex",alignItems:"center",justifyContent:"center",
                fontSize:12,fontWeight:800,color:HW.white}}>{i+1}</div>
              <span style={{fontSize:12,color:HW.muted}}>{step}</span>
              {i<2&&<div style={{width:30,height:1,background:HW.border}}/>}
            </div>
          ))}
        </div>
        <div style={s.card}>
          <h3 style={{marginBottom:20,color:HW.red}}>👤 Personal Information</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:20}}>
            <div style={{gridColumn:"1/-1"}}>
              <label style={s.label}>Full Name *</label>
              <input style={s.input} placeholder="e.g. Ahmed Mohammed Al-Rashidi"
                value={setupProfile.full_name}
                onChange={e=>setSetupProfile({...setupProfile,full_name:e.target.value})}/></div>
            <div><label style={s.label}>Civil ID *</label>
              <input style={s.input} placeholder="e.g. 10234567"
                value={setupProfile.civil_id}
                onChange={e=>setSetupProfile({...setupProfile,civil_id:e.target.value})}/></div>
            <div><label style={s.label}>Date of Birth</label>
              <input style={s.input} type="date" value={setupProfile.date_of_birth}
                onChange={e=>setSetupProfile({...setupProfile,date_of_birth:e.target.value})}/></div>
            <div><label style={s.label}>Gender</label>
              <select style={s.input} value={setupProfile.gender}
                onChange={e=>setSetupProfile({...setupProfile,gender:e.target.value})}>
                <option value="">Select…</option>
                <option value="Male">Male / ذكر</option>
                <option value="Female">Female / أنثى</option>
              </select></div>
            <div><label style={s.label}>Nationality</label>
              <input style={s.input} placeholder="e.g. Omani"
                value={setupProfile.nationality}
                onChange={e=>setSetupProfile({...setupProfile,nationality:e.target.value})}/></div>
          </div>
          <h3 style={{marginBottom:16,color:HW.red,paddingTop:16,
            borderTop:`1px solid ${HW.border}`}}>📚 Academic Information</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:20}}>
            <div><label style={s.label}>Department</label>
              <input style={s.input} placeholder="e.g. Engineering"
                value={setupProfile.department}
                onChange={e=>setSetupProfile({...setupProfile,department:e.target.value})}/></div>
            <div><label style={s.label}>GPA</label>
              <input style={s.input} type="number" placeholder="e.g. 3.85"
                min="0" max="4" step="0.01" value={setupProfile.gpa}
                onChange={e=>setSetupProfile({...setupProfile,gpa:e.target.value})}/></div>
            <div style={{gridColumn:"1/-1"}}><label style={s.label}>Assigned Mentor</label>
              <input style={s.input} placeholder="e.g. Dr. Fatima Al-Sayed"
                value={setupProfile.assigned_mentor}
                onChange={e=>setSetupProfile({...setupProfile,assigned_mentor:e.target.value})}/></div>
          </div>
          <h3 style={{marginBottom:16,color:HW.red,paddingTop:16,
            borderTop:`1px solid ${HW.border}`}}>📞 Contact Information</h3>
          <div><label style={s.label}>Phone Number</label>
            <input style={s.input} placeholder="e.g. +968-9100-0001"
              value={setupProfile.phone_number}
              onChange={e=>setSetupProfile({...setupProfile,phone_number:e.target.value})}/></div>
          <button style={{...s.btn,background:HW.red,color:HW.white,
            width:"100%",padding:14,fontSize:15,marginTop:24,opacity:loading?0.6:1}}
            onClick={saveSetupProfile} disabled={loading}>
            {loading?"Saving your profile…":"Complete Profile & Start →"}
          </button>
          {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
            fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
        </div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════
  // TRAINEE DASHBOARD
  // ══════════════════════════════════════════════════
  if(view==="trainee") return (
    <div style={s.page}>
      {showReminder&&<ReminderPopup onDismiss={()=>setShowReminder(false)}/>}
      {showGreeting&&<GreetingPopup name={traineeName} onDismiss={()=>setShowGreeting(false)}/>}

      {/* Topbar */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
        marginBottom:24,paddingBottom:20,borderBottom:`1px solid ${HW.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <HuaweiLogo size={36}/>
          <div>
            <h2 style={{margin:0,fontSize:20}}>
              {traineeName?`Welcome, ${traineeName}!`:"Trainee Dashboard"}
            </h2>
            <p style={{color:HW.muted,fontSize:12,margin:0}}>{user?.email}</p>
          </div>
        </div>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:28,fontWeight:800,color:clockColor,
            fontFamily:"monospace",letterSpacing:3}}>{hh}:{mm}:{ss}</div>
          <div style={{fontSize:11,color:clockColor,fontWeight:700}}>
            {isLate?"⚠️ Past 9:00 AM":"✅ On Time"}
          </div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:12}}
            onClick={()=>setShowReminder(true)}>⏰</button>
          <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:12}}
            onClick={()=>setShowGreeting(true)}>👋</button>
          <button style={{...s.btn,background:HW.red,color:HW.white}}
            onClick={logout}>Sign out</button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",gap:4,background:HW.surface2,
        borderRadius:12,padding:5,marginBottom:28}}>
        {[
          {id:"attendance",icon:"✅",label:"Attendance"},
          {id:"tasks",icon:"📋",label:"Daily Tasks"},
          {id:"goals",icon:"🎯",label:"My Goals"},
        ].map(tab=>(
          <button key={tab.id} onClick={()=>setTraineeTab(tab.id)}
            style={{...s.btn,flex:1,padding:"13px",
              background:traineeTab===tab.id?HW.red:"none",
              color:traineeTab===tab.id?HW.white:HW.muted,
              fontSize:14,borderRadius:8,fontWeight:700,transition:"all .2s"}}>
            {tab.icon} {tab.label}
          </button>
        ))}
      </div>

      {/* ══ ATTENDANCE TAB ══ */}
      {traineeTab==="attendance"&&(
        <div>
          {/* Big clock */}
          <div style={{...s.card,textAlign:"center",border:`1px solid ${clockColor}40`}}>
            <div style={{fontSize:11,color:HW.muted,fontWeight:700,
              textTransform:"uppercase",letterSpacing:".1em",marginBottom:8}}>
              Current Time
            </div>
            <div style={{fontSize:64,fontWeight:800,color:clockColor,
              fontFamily:"monospace",letterSpacing:8,
              textShadow:`0 0 40px ${clockColor}40`}}>
              {hh}:{mm}:{ss}
            </div>
            <div style={{fontSize:13,color:HW.muted,marginTop:8}}>
              {currentTime.toLocaleDateString("en-GB",{
                weekday:"long",day:"numeric",month:"long",year:"numeric"})}
            </div>
            <div style={{marginTop:12,display:"inline-block",
              padding:"6px 20px",borderRadius:20,fontWeight:700,fontSize:13,
              background:`${clockColor}15`,color:clockColor,
              border:`1px solid ${clockColor}40`}}>
              {isLate?"⚠️ Past 9:00 AM — Excuse Required":"✅ On Time — You Can Sign In"}
            </div>
          </div>

          {isLate&&(
            <div style={{background:"rgba(207,10,44,.08)",
              border:`1px solid rgba(207,10,44,.3)`,
              borderRadius:12,padding:16,marginBottom:20,
              display:"flex",alignItems:"center",gap:12}}>
              <span style={{fontSize:28}}>🔒</span>
              <div>
                <div style={{fontWeight:700,color:HW.red,fontSize:15}}>
                  Attendance Time Locked
                </div>
                <div style={{fontSize:13,color:HW.muted,marginTop:4}}>
                  Sign-in deadline was <b style={{color:HW.text}}>9:00 AM</b>.
                  Provide an excuse below. A{" "}
                  <b style={{color:HW.red}}>penalty of 8.33%</b> will apply.
                </div>
              </div>
            </div>
          )}

          <div style={s.card}>
            <h3 style={{marginBottom:18,color:HW.red}}>✅ Sign Attendance</h3>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <div><label style={s.label}>Date</label>
                <input style={s.input} type="date" value={attendance.report_date}
                  onChange={e=>setAttendance({...attendance,report_date:e.target.value})}/></div>
              <div>
                <label style={s.label}>Your Sign-in Time (auto)</label>
                <div style={{background:HW.surface2,border:`1px solid ${HW.border}`,
                  borderRadius:8,padding:"9px 13px",fontSize:16,
                  color:clockColor,fontWeight:800,fontFamily:"monospace"}}>
                  {hh}:{mm}:{ss}
                </div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10,paddingTop:22}}>
                <input type="checkbox" id="att" checked={attendance.attended}
                  onChange={e=>setAttendance({...attendance,attended:e.target.checked})}/>
                <label htmlFor="att" style={{...s.label,margin:0}}>
                  I confirm attendance today
                </label>
              </div>
              {!attendance.attended&&(
                <div style={{background:"rgba(207,10,44,.1)",
                  border:`1px solid rgba(207,10,44,.3)`,
                  borderRadius:8,padding:10,fontSize:12,color:HW.red}}>
                  ⚠️ <b>Penalty Warning:</b> Not marked as attended — <b>8.33% deduction</b>
                </div>
              )}
            </div>

            {isLate&&(
              <div style={{marginTop:16,background:HW.surface2,borderRadius:12,
                padding:16,border:`1px solid rgba(207,10,44,.3)`}}>
                <div style={{fontWeight:700,color:HW.red,marginBottom:12}}>
                  ⚠️ Late Arrival — Excuse Required
                </div>
                <div style={{marginBottom:10}}>
                  <label style={s.label}>Excuse Type</label>
                  <select style={s.input} value={excuseType}
                    onChange={e=>setExcuseType(e.target.value)}>
                    <option value="">Select reason…</option>
                    <option value="traffic">🚗 Road Traffic {trafficCount>=2?"(LIMIT REACHED)":""}</option>
                    <option value="medical">🏥 Medical Emergency</option>
                    <option value="family">👨‍👩‍👧 Family Emergency</option>
                    <option value="other">📋 Other</option>
                  </select>
                </div>
                <div style={{marginBottom:10}}>
                  <label style={s.label}>Excuse Description *</label>
                  <textarea style={{...s.input,height:80,resize:"vertical"}}
                    placeholder="Please describe your reason…"
                    value={excuseText} onChange={e=>setExcuseText(e.target.value)}/>
                </div>
                <div>
                  <label style={s.label}>
                    Proof Photo {excuseType==="traffic"?"(Required)":"(Optional)"}
                  </label>
                  <div style={{display:"flex",gap:10,alignItems:"center"}}>
                    <button style={{...s.btn,background:HW.surface,color:HW.text,
                      border:`1px dashed ${HW.border}`}}
                      onClick={()=>excuseRef.current.click()}>
                      📷 {excusePhoto?"Change":"Upload Proof"}
                    </button>
                    <input ref={excuseRef} type="file" accept="image/*"
                      style={{display:"none"}} onChange={handleExcusePhoto}/>
                    {excusePreview&&(
                      <img src={excusePreview} alt="proof"
                        style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>
                    )}
                  </div>
                  {excuseType==="traffic"&&(
                    <p style={{fontSize:11,color:HW.muted,marginTop:6}}>
                      Traffic excuses this month:{" "}
                      <b style={{color:trafficCount>=2?HW.red:HW.text}}>{trafficCount}/2</b>
                    </p>
                  )}
                </div>
              </div>
            )}

            {/* Face ID */}
            <div style={{marginTop:16,padding:16,background:HW.surface2,borderRadius:12,
              border:faceCaptured?`1px solid rgba(207,10,44,.5)`:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:12}}>🤳 Face ID Verification</div>
              {!cameraActive&&!faceCaptured&&(
                <button style={{...s.btn,background:HW.red,color:HW.white}}
                  onClick={startCamera}>📷 Open Camera</button>
              )}
              {cameraActive&&(
                <div>
                  <video ref={videoRef} autoPlay playsInline
                    style={{width:"100%",maxWidth:280,borderRadius:10,
                      border:`2px solid ${HW.red}`,display:"block",marginBottom:10}}/>
                  <canvas ref={canvasRef} style={{display:"none"}}/>
                  <div style={{display:"flex",gap:8}}>
                    <button style={{...s.btn,background:HW.red,color:HW.white}}
                      onClick={captureface}>📸 Capture</button>
                    <button style={{...s.btn,background:HW.surface,color:HW.muted,
                      border:`1px solid ${HW.border}`}} onClick={stopCamera}>Cancel</button>
                  </div>
                </div>
              )}
              {faceCaptured&&facePreviewUrl&&(
                <div style={{display:"flex",alignItems:"center",gap:16}}>
                  <img src={facePreviewUrl} alt="Face ID"
                    style={{width:90,height:90,borderRadius:"50%",objectFit:"cover",
                      border:`3px solid ${HW.red}`,
                      boxShadow:`0 0 20px rgba(207,10,44,.3)`}}/>
                  <div>
                    <div style={{fontWeight:700,color:HW.red,fontSize:15,marginBottom:4}}>
                      ✅ Face Captured Successfully
                    </div>
                    <div style={{fontSize:12,color:HW.muted,marginBottom:8}}>
                      Identity verified
                    </div>
                    <button style={{...s.btn,background:"rgba(207,10,44,.1)",
                      color:HW.red,padding:"5px 12px",fontSize:12}}
                      onClick={()=>{
                        setFaceCaptured(false);setFaceCapture(null);setFacePreviewUrl(null);
                      }}>🔄 Retake</button>
                  </div>
                </div>
              )}
            </div>

            {/* Location */}
            <div style={{marginTop:12,padding:16,background:HW.surface2,borderRadius:12,
              border:locationOk?`1px solid rgba(207,10,44,.5)`:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:10}}>📍 Location Verification</div>
              <p style={{fontSize:12,color:HW.muted,marginBottom:10}}>
                Must be within <b style={{color:HW.text}}>{WORK_RADIUS}m</b> of workplace.
              </p>
              <button style={{...s.btn,
                background:locationOk?"rgba(207,10,44,.15)":HW.red,
                color:locationOk?HW.red:HW.white,opacity:geoStatus==="checking"?0.6:1}}
                onClick={checkLocation} disabled={geoStatus==="checking"}>
                {geoStatus==="checking"?"Checking…":locationOk?"✅ Verified":"📍 Verify Location"}
              </button>
              {geoMsg&&<p style={{fontSize:12,marginTop:8,
                color:geoStatus==="ok"?"#34d399":HW.red}}>{geoMsg}</p>}
            </div>

            {/* Submit Attendance Button */}
            <button style={{...s.btn,
              background:(locationOk&&faceCaptured)?HW.red:HW.surface2,
              color:(locationOk&&faceCaptured)?HW.white:HW.muted,
              marginTop:16,width:"100%",padding:14,fontSize:15,
              opacity:loading?0.6:1,
              cursor:(locationOk&&faceCaptured)?"pointer":"not-allowed"}}
              onClick={submitAttendance}
              disabled={loading||!locationOk||!faceCaptured}>
              {loading?"Saving…":"✅ Submit Attendance"}
            </button>

            {/* Sign Out Button */}
            <div style={{marginTop:10,padding:16,background:HW.surface2,
              borderRadius:12,border:`1px solid ${HW.border}`}}>
              <div style={{fontWeight:700,color:"#4f8ef7",marginBottom:8}}>
                🚪 Sign Out
              </div>
              {signedOut?(
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:40,height:40,borderRadius:"50%",
                    background:"rgba(79,142,247,.15)",display:"flex",
                    alignItems:"center",justifyContent:"center",fontSize:20}}>✅</div>
                  <div>
                    <div style={{fontWeight:700,color:"#4f8ef7"}}>
                      Signed out at {signoutTime}
                    </div>
                    <div style={{fontSize:12,color:HW.muted}}>
                      Your time out has been recorded
                    </div>
                  </div>
                </div>
              ):(
                <button style={{...s.btn,
                  background:"rgba(79,142,247,.15)",color:"#4f8ef7",
                  width:"100%",padding:12,fontSize:14,
                  border:`1px solid rgba(79,142,247,.3)`,
                  opacity:loading?0.6:1}}
                  onClick={submitSignOut} disabled={loading}>
                  {loading?"Recording…":"🚪 Record Time Out"}
                </button>
              )}
            </div>

            {(!locationOk||!faceCaptured)&&(
              <p style={{fontSize:12,color:HW.muted,textAlign:"center",marginTop:8}}>
                {!locationOk&&!faceCaptured?"Verify location and face ID above"
                  :!locationOk?"📍 Verify location first":"🤳 Complete face ID first"}
              </p>
            )}
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
        </div>
      )}

      {/* ══ TASKS TAB ══ */}
      {traineeTab==="tasks"&&(
        <div>
          <div style={s.card}>
            <h3 style={{marginBottom:8,color:HW.red}}>📋 Daily Tasks Report</h3>
            <p style={{color:HW.muted,fontSize:13,marginBottom:20}}>
              Describe everything you worked on today. AI will analyze your report and generate a KPI score.
            </p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <div><label style={s.label}>Report Date</label>
                <input style={s.input} type="date" value={taskReport.report_date}
                  onChange={e=>setTaskReport({...taskReport,report_date:e.target.value})}/></div>
              <div style={{display:"flex",alignItems:"center",gap:8,paddingTop:22}}>
                <span style={{fontSize:20}}>💡</span>
                <span style={{fontSize:12,color:HW.muted}}>More detail = higher AI score!</span>
              </div>
              <div style={{gridColumn:"1/-1"}}>
                <label style={s.label}>What did you work on today? *</label>
                <textarea style={{...s.input,height:180,resize:"vertical"}}
                  placeholder="Example: Completed the circuit analysis report and submitted it to my mentor. Attended a 2-hour lab session. Reviewed safety protocols..."
                  value={taskReport.report_text}
                  onChange={e=>setTaskReport({...taskReport,report_text:e.target.value})}/>
              </div>
              <div style={{gridColumn:"1/-1"}}>
                <label style={s.label}>📸 Proof Photo (optional)</label>
                <div style={{display:"flex",gap:12,alignItems:"center"}}>
                  <button style={{...s.btn,background:HW.surface2,color:HW.text,
                    border:`1px dashed ${HW.border}`}}
                    onClick={()=>photoRef.current.click()}>
                    {photoFile?"📷 Change Photo":"📷 Upload Photo"}
                  </button>
                  <input ref={photoRef} type="file" accept="image/*"
                    style={{display:"none"}} onChange={handleProofPhoto}/>
                  {photoPreview&&(
                    <img src={photoPreview} alt="preview"
                      style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>
                  )}
                </div>
              </div>
            </div>
            <button style={{...s.btn,background:HW.red,color:HW.white,
              marginTop:16,width:"100%",padding:14,fontSize:15,
              opacity:(loading||aiLoading)?0.6:1}}
              onClick={submitTasks} disabled={loading||aiLoading}>
              {aiLoading?"🤖 AI Analyzing…":loading?"Saving…":"📋 Submit Daily Tasks ✓"}
            </button>
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":
              msg.startsWith("🤖")?"#FFA500":HW.red,
              fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
          {aiResult&&(
            <div style={{...s.card,border:`1px solid rgba(207,10,44,.4)`}}>
              <h3 style={{marginBottom:20,color:HW.red}}>🤖 AI Analysis Result</h3>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
                <div style={{background:HW.surface2,borderRadius:12,padding:16,textAlign:"center"}}>
                  <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                    textTransform:"uppercase",marginBottom:8}}>KPI Score</div>
                  <div style={{fontSize:52,fontWeight:800,color:kpiColor(aiResult.kpi_score)}}>
                    {aiResult.kpi_score}
                  </div>
                  <div style={{fontSize:11,color:HW.muted}}>out of 100</div>
                </div>
                <div style={{background:HW.surface2,borderRadius:12,padding:16}}>
                  <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                    textTransform:"uppercase",marginBottom:12}}>Task Breakdown</div>
                  {aiResult.pie_chart&&<PieChart data={aiResult.pie_chart}/>}
                </div>
              </div>
              {aiResult.summary&&(
                <div style={{background:HW.surface2,borderRadius:10,padding:14,
                  marginBottom:12,borderLeft:`3px solid ${HW.red}`}}>
                  <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                    marginBottom:6,textTransform:"uppercase"}}>Day Summary</div>
                  <div style={{fontSize:14}}>{aiResult.summary}</div>
                </div>
              )}
              {aiResult.talent_notes&&(
                <div style={{background:HW.surface2,borderRadius:10,padding:14,
                  borderLeft:"3px solid #FFA500"}}>
                  <div style={{fontSize:11,color:"#FFA500",fontWeight:700,
                    marginBottom:6,textTransform:"uppercase"}}>🌟 Talent Notes</div>
                  <div style={{fontSize:14,lineHeight:1.6}}>{aiResult.talent_notes}</div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* ══ GOALS TAB ══ */}
      {traineeTab==="goals"&&(
        <div>
          <div style={{display:"grid",
            gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",
            gap:12,marginBottom:24}}>
            {KRA_CATEGORIES.map(kra=>{
              const kraGoals=goals.filter(g=>g.kra===kra.id);
              const completed=kraGoals.filter(g=>g.status==="completed").length;
              return (
                <div key={kra.id} style={{background:HW.surface,
                  border:`1px solid ${HW.border}`,borderRadius:12,padding:16,
                  cursor:"pointer",borderTop:`3px solid ${kra.color}`,
                  opacity:goalFilter===kra.id?1:0.7,
                  outline:goalFilter===kra.id?`2px solid ${kra.color}`:"none"}}
                  onClick={()=>setGoalFilter(goalFilter===kra.id?"all":kra.id)}>
                  <div style={{fontSize:22,marginBottom:6}}>{kra.icon}</div>
                  <div style={{fontSize:11,fontWeight:700,color:kra.color,
                    textTransform:"uppercase",marginBottom:4}}>{kra.label}</div>
                  <div style={{fontSize:13,color:HW.muted}}>
                    {kraGoals.length} goals · {completed} done
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",
            gap:12,marginBottom:20}}>
            {[
              {label:"Total Goals",value:goals.length,color:HW.red},
              {label:"In Progress",value:goals.filter(g=>g.status==="in_progress").length,color:"#4f8ef7"},
              {label:"Completed",value:goals.filter(g=>g.status==="completed").length,color:"#34d399"},
              {label:"Overdue",value:goals.filter(g=>g.due_date&&
                new Date(g.due_date)<new Date()&&g.status!=="completed").length,color:"#f87171"},
            ].map((stat,i)=>(
              <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                borderRadius:12,padding:14,textAlign:"center",
                borderTop:`3px solid ${stat.color}`}}>
                <div style={{fontSize:24,fontWeight:800,color:stat.color}}>{stat.value}</div>
                <div style={{fontSize:11,color:HW.muted,marginTop:4}}>{stat.label}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",justifyContent:"space-between",
            alignItems:"center",marginBottom:16}}>
            <h3 style={{margin:0}}>
              {goalFilter==="all"?"All Goals"
                :`${KRA_CATEGORIES.find(k=>k.id===goalFilter)?.label||""} Goals`}
              <span style={{fontSize:13,color:HW.muted,fontWeight:400,marginLeft:8}}>
                ({filteredGoals.length})
              </span>
            </h3>
            <div style={{display:"flex",gap:8}}>
              {goalFilter!=="all"&&(
                <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:12}}
                  onClick={()=>setGoalFilter("all")}>Show All</button>
              )}
              <button style={{...s.btn,background:HW.red,color:HW.white}}
                onClick={()=>setShowAddGoal(!showAddGoal)}>
                {showAddGoal?"Cancel":"+ Add Goal"}
              </button>
            </div>
          </div>
          {showAddGoal&&(
            <div style={{...s.card,border:`1px solid rgba(207,10,44,.3)`,marginBottom:20}}>
              <h4 style={{marginBottom:16,color:HW.red}}>🎯 Set New Goal</h4>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                <div><label style={s.label}>KRA Category</label>
                  <select style={s.input} value={newGoal.kra}
                    onChange={e=>setNewGoal({...newGoal,kra:e.target.value})}>
                    {KRA_CATEGORIES.map(k=>(
                      <option key={k.id} value={k.id}>{k.icon} {k.label}</option>
                    ))}
                  </select></div>
                <div><label style={s.label}>Status</label>
                  <select style={s.input} value={newGoal.status}
                    onChange={e=>setNewGoal({...newGoal,status:e.target.value})}>
                    <option value="not_started">⬜ Not Started</option>
                    <option value="in_progress">🔵 In Progress</option>
                  </select></div>
                <div style={{gridColumn:"1/-1"}}><label style={s.label}>Goal Title *</label>
                  <input style={s.input}
                    placeholder="e.g. Achieve 95% attendance this month"
                    value={newGoal.goal_title}
                    onChange={e=>setNewGoal({...newGoal,goal_title:e.target.value})}/></div>
                <div style={{gridColumn:"1/-1"}}><label style={s.label}>Description</label>
                  <textarea style={{...s.input,height:70,resize:"vertical"}}
                    placeholder="How do you plan to achieve this?"
                    value={newGoal.description}
                    onChange={e=>setNewGoal({...newGoal,description:e.target.value})}/></div>
                <div><label style={s.label}>Target Value</label>
                  <input style={s.input} type="number" value={newGoal.target_value}
                    onChange={e=>setNewGoal({...newGoal,target_value:parseFloat(e.target.value)})}/></div>
                <div><label style={s.label}>Unit</label>
                  <select style={s.input} value={newGoal.unit}
                    onChange={e=>setNewGoal({...newGoal,unit:e.target.value})}>
                    <option value="%">%</option>
                    <option value="days">Days</option>
                    <option value="sessions">Sessions</option>
                    <option value="reports">Reports</option>
                    <option value="tasks">Tasks</option>
                    <option value="hours">Hours</option>
                    <option value="score">Score</option>
                  </select></div>
                <div><label style={s.label}>Start Date</label>
                  <input style={s.input} type="date" value={newGoal.start_date}
                    onChange={e=>setNewGoal({...newGoal,start_date:e.target.value})}/></div>
                <div><label style={s.label}>Due Date</label>
                  <input style={s.input} type="date" value={newGoal.due_date}
                    onChange={e=>setNewGoal({...newGoal,due_date:e.target.value})}/></div>
              </div>
              <button style={{...s.btn,background:HW.red,color:HW.white,
                marginTop:16,opacity:loading?0.6:1}}
                onClick={addGoal} disabled={loading}>
                {loading?"Saving…":"🎯 Set Goal"}
              </button>
              {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                fontSize:13,marginTop:10}}>{msg}</p>}
            </div>
          )}
          {filteredGoals.length===0?(
            <div style={{...s.card,textAlign:"center",padding:40}}>
              <div style={{fontSize:40,marginBottom:12}}>🎯</div>
              <p style={{color:HW.muted}}>No goals yet. Click "+ Add Goal" to get started!</p>
            </div>
          ):(
            filteredGoals.map(goal=>(
              <GoalCard key={goal.id} goal={goal}
                onUpdate={updateGoal} onDelete={deleteGoal} isTrainee={true}/>
            ))
          )}
          {msg&&!showAddGoal&&(
            <p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>
          )}
        </div>
      )}
    </div>
  );

  // ══════════════════════════════════════════════════
  // MANAGEMENT DASHBOARD
  // ══════════════════════════════════════════════════
  if(view==="mgmt") {
    const analytics=getAnalytics();
    return (
      <div style={s.page}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
          marginBottom:28,paddingBottom:20,borderBottom:`1px solid ${HW.border}`}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <HuaweiLogo size={36}/>
            <div>
              <h2 style={{margin:0,fontSize:20}}>Management Dashboard</h2>
              <p style={{color:HW.muted,fontSize:12,margin:0}}>{user?.email}</p>
            </div>
          </div>
          <div style={{textAlign:"center"}}>
            <div style={{fontSize:28,fontWeight:800,color:clockColor,
              fontFamily:"monospace",letterSpacing:3}}>{hh}:{mm}:{ss}</div>
            <div style={{fontSize:11,color:clockColor,fontWeight:700}}>
              {isLate?"⚠️ Past Sign-in Deadline":"✅ Sign-in Window Open"}
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button style={{...s.btn,background:"rgba(207,10,44,.15)",color:HW.red}}
              onClick={exportExcel}>📊 Export Excel</button>
            <button style={{...s.btn,background:HW.red,color:HW.white}}
              onClick={logout}>Sign out</button>
          </div>
        </div>

        <div style={{display:"flex",gap:4,background:HW.surface2,
          borderRadius:10,padding:4,marginBottom:24}}>
          {["trainees","live","analytics","okr"].map(tab=>(
            <button key={tab}
              onClick={()=>{setMgmtTab(tab);setSelected(null);setMsg("");}}
              style={{...s.btn,flex:1,padding:"10px",
                background:mgmtTab===tab?HW.surface:"none",
                color:mgmtTab===tab?HW.text:HW.muted,fontSize:13,borderRadius:7,
                borderBottom:mgmtTab===tab?`2px solid ${HW.red}`:"none"}}>
              {tab==="trainees"?"👥 Trainees":tab==="live"?"📡 Live Sign-ins":
               tab==="analytics"?"📊 KPI Analytics":"🎯 OKR"}
            </button>
          ))}
        </div>

        {/* LIVE */}
        {mgmtTab==="live"&&(
          <div>
            <div style={{display:"grid",
              gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",
              gap:16,marginBottom:24}}>
              {[
                {label:"Signed In Today",value:liveSignins.filter(s=>s.attended).length,icon:"✅",color:"#34d399"},
                {label:"Absent Today",value:liveSignins.filter(s=>!s.attended).length,icon:"❌",color:HW.red},
                {label:"On Time",value:liveSignins.filter(s=>s.signin_time&&s.signin_time<=MAX_SIGNIN).length,icon:"⏰",color:"#4f8ef7"},
                {label:"Late",value:liveSignins.filter(s=>s.signin_time&&s.signin_time>MAX_SIGNIN).length,icon:"⚠️",color:"#FFA500"},
              ].map((stat,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:14,padding:20,borderTop:`3px solid ${stat.color}`}}>
                  <div style={{fontSize:24,marginBottom:8}}>{stat.icon}</div>
                  <div style={{fontSize:28,fontWeight:800,color:stat.color}}>{stat.value}</div>
                  <div style={{fontSize:12,color:HW.muted,marginTop:4}}>{stat.label}</div>
                </div>
              ))}
            </div>
            <div style={s.card}>
              <div style={{display:"flex",justifyContent:"space-between",
                alignItems:"center",marginBottom:16}}>
                <h3 style={{margin:0}}>📡 Live Sign-in Feed</h3>
                <div style={{display:"flex",alignItems:"center",gap:6,
                  background:"rgba(52,211,153,.1)",border:"1px solid rgba(52,211,153,.3)",
                  borderRadius:20,padding:"4px 12px"}}>
                  <div style={{width:8,height:8,borderRadius:"50%",
                    background:"#34d399",animation:"pulse 1.5s infinite"}}/>
                  <span style={{fontSize:12,color:"#34d399",fontWeight:700}}>LIVE</span>
                </div>
              </div>
              <style>{`
                @keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(1.3)}}
                @keyframes slideIn{from{transform:translateX(-20px);opacity:0}to{transform:translateX(0);opacity:1}}
              `}</style>
              {liveSignins.length===0?(
                <div style={{textAlign:"center",padding:40}}>
                  <div style={{fontSize:40,marginBottom:12}}>📡</div>
                  <p style={{color:HW.muted}}>No sign-ins yet today. Waiting…</p>
                </div>
              ):(
                liveSignins.map((signin,i)=>(
                  <div key={signin.id} style={{display:"flex",alignItems:"center",
                    gap:16,padding:"14px 16px",background:HW.surface2,
                    borderRadius:12,marginBottom:10,
                    borderLeft:`4px solid ${signin.attended?"#34d399":HW.red}`,
                    animation:i===0?"slideIn .4s ease":"none"}}>
                    <div style={{width:44,height:44,borderRadius:"50%",
                      background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:16,fontWeight:800,color:HW.white,flexShrink:0}}>
                      {signin.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}
                    </div>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:700,fontSize:15}}>{signin.full_name}</div>
                      <div style={{fontSize:12,color:HW.muted,marginTop:2}}>{signin.department}</div>
                    </div>
                    <div style={{textAlign:"center"}}>
                      <div style={{fontSize:20,fontWeight:800,fontFamily:"monospace",
                        color:signin.signin_time>MAX_SIGNIN?HW.red:"#34d399"}}>
                        {signin.signin_time||"—"}
                      </div>
                      <div style={{fontSize:10,color:HW.muted}}>Sign-in</div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <span style={{padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700,
                        background:signin.attended?"rgba(52,211,153,.15)":"rgba(207,10,44,.15)",
                        color:signin.attended?"#34d399":HW.red}}>
                        {signin.attended?"● Present":"○ Absent"}
                      </span>
                      {signin.penalty_applied&&(
                        <div style={{fontSize:11,color:HW.red,marginTop:4,fontWeight:700}}>
                          ⚠️ Penalty
                        </div>
                      )}
                      <div style={{fontSize:10,color:HW.muted,marginTop:4}}>
                        {signin.timestamp instanceof Date
                          ?signin.timestamp.toLocaleTimeString("en-GB",
                            {hour:"2-digit",minute:"2-digit"}):"—"}
                      </div>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )}

        {/* TRAINEES */}
        {mgmtTab==="trainees"&&!selected&&(
          <div style={s.card}>
            <h3 style={{marginBottom:16}}>All Trainees</h3>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {["Name","Department","Mentor","GPA","Joined","Status",""].map(h=>(
                  <th key={h} style={s.th}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {trainees.map(t=>(
                  <tr key={t.id} style={{cursor:"pointer",opacity:t.status==="dropped"?0.4:1}}
                    onClick={()=>openProfile(t)}>
                    <td style={s.td}><b>{t.full_name}</b></td>
                    <td style={s.td}>{t.department}</td>
                    <td style={{...s.td,color:HW.muted,fontSize:13}}>{t.assigned_mentor}</td>
                    <td style={s.td}><b>{t.gpa}</b></td>
                    <td style={{...s.td,color:HW.muted,fontSize:12}}>{t.joining_date||"—"}</td>
                    <td style={s.td}>
                      <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,
                        background:statusColors[t.status]?.bg,color:statusColors[t.status]?.color}}>
                        {t.status}
                      </span>
                    </td>
                    <td style={s.td}>
                      <button style={{...s.btn,background:"rgba(207,10,44,.1)",color:HW.red,
                        fontSize:11,padding:"4px 10px"}}
                        onClick={e=>{e.stopPropagation();exportPDF(t);}}>📄 PDF</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {trainees.length===0&&(
              <div style={{textAlign:"center",padding:40}}>
                <div style={{fontSize:40,marginBottom:12}}>👥</div>
                <p style={{color:HW.muted}}>No trainees yet.</p>
              </div>
            )}
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
        )}

        {mgmtTab==="trainees"&&selected&&(
          <div>
            <button style={{...s.btn,background:HW.surface2,color:HW.text,
              marginBottom:20,border:`1px solid ${HW.border}`}}
              onClick={()=>{setSelected(null);setMsg("");}}>← Back to table</button>
            <div style={{...s.card,display:"flex",alignItems:"center",
              justifyContent:"space-between",flexWrap:"wrap",gap:16}}>
              <div style={{display:"flex",alignItems:"center",gap:16}}>
                <div style={{width:56,height:56,borderRadius:14,
                  background:`linear-gradient(135deg,${HW.red},${HW.darkRed})`,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:22,fontWeight:800,color:HW.white}}>
                  {selected.full_name?.split(" ").map(w=>w[0]).join("").slice(0,2)}
                </div>
                <div>
                  <div style={{fontWeight:700,fontSize:18}}>{selected.full_name}</div>
                  <div style={{color:HW.muted,fontSize:13}}>
                    {selected.department} · {selected.assigned_mentor}
                  </div>
                  <div style={{display:"flex",gap:16,marginTop:4,flexWrap:"wrap"}}>
                    <span style={{fontSize:12,color:HW.muted}}>
                      📅 Joined: <b style={{color:HW.text}}>{selected.joining_date||"—"}</b>
                    </span>
                    {selected.quitting_date&&(
                      <span style={{fontSize:12,color:HW.muted}}>
                        🚪 Quit: <b style={{color:HW.red}}>{selected.quitting_date}</b>
                      </span>
                    )}
                    {selected.joining_date&&(
                      <span style={{fontSize:12,color:HW.muted}}>
                        ⏱ <b style={{color:HW.text}}>
                          {Math.floor((new Date(selected.quitting_date||new Date())-
                            new Date(selected.joining_date))/(1000*60*60*24))} days
                        </b>
                      </span>
                    )}
                  </div>
                  <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,
                    marginTop:6,display:"inline-block",
                    background:statusColors[selected.status]?.bg,
                    color:statusColors[selected.status]?.color}}>
                    {selected.status}
                  </span>
                </div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button style={{...s.btn,background:"rgba(207,10,44,.15)",color:HW.red,fontSize:12}}
                  onClick={()=>exportPDF()}>📄 PDF</button>
                {selected.status!=="active"&&(
                  <button style={{...s.btn,background:"rgba(207,10,44,.15)",color:HW.red,fontSize:12}}
                    onClick={()=>changeStatus("active")}>✅ Reactivate</button>
                )}
                {selected.status!=="transferred"&&(
                  <button style={{...s.btn,background:"rgba(255,165,0,.15)",color:"#FFA500",fontSize:12}}
                    onClick={()=>changeStatus("transferred")}>🔄 Transferred</button>
                )}
                {selected.status!=="inactive"&&(
                  <button style={{...s.btn,background:"rgba(136,136,136,.15)",color:HW.muted,fontSize:12}}
                    onClick={()=>changeStatus("inactive")}>⏸ Inactive</button>
                )}
                {selected.status!=="dropped"&&(
                  <button style={{...s.btn,background:"rgba(100,100,100,.2)",color:"#666",fontSize:12}}
                    onClick={()=>changeStatus("dropped")}>🔴 Dropped</button>
                )}
              </div>
            </div>

            <div style={{display:"flex",gap:4,background:HW.surface2,
              borderRadius:10,padding:4,marginBottom:20}}>
              {["timeline","edit","reports","penalties","goals"].map(tab=>(
                <button key={tab} onClick={()=>setProfileTab(tab)}
                  style={{...s.btn,flex:1,padding:"8px",
                    background:profileTab===tab?HW.surface:"none",
                    color:profileTab===tab?HW.text:HW.muted,fontSize:11,borderRadius:7,
                    borderBottom:profileTab===tab?`2px solid ${HW.red}`:"none"}}>
                  {tab==="timeline"?"📅 Timeline":tab==="edit"?"✏️ Edit":
                   tab==="reports"?"📋 Reports":tab==="penalties"?"⚠️ Penalties":"🎯 Goals"}
                </button>
              ))}
            </div>

            {profileTab==="timeline"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:20}}>Activity Timeline</h3>
                {logs.length===0?<p style={{color:HW.muted}}>No activity yet.</p>
                  :logs.map((log,i)=>(
                    <div key={log.id} style={{display:"flex",gap:16,
                      marginBottom:20,position:"relative"}}>
                      {i<logs.length-1&&(
                        <div style={{position:"absolute",left:19,top:40,width:2,
                          height:"calc(100% + 4px)",background:HW.border}}/>
                      )}
                      <div style={{width:40,height:40,borderRadius:"50%",
                        background:HW.surface2,border:`2px solid ${HW.border}`,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        fontSize:18,flexShrink:0,zIndex:1}}>
                        {eventIcons[log.event_type]||"📌"}
                      </div>
                      <div style={{flex:1,paddingTop:6}}>
                        <div style={{fontWeight:600,fontSize:14}}>{log.description}</div>
                        {log.old_value&&log.new_value&&(
                          <div style={{fontSize:12,color:HW.muted,marginTop:4}}>
                            <span style={{color:"#f87171"}}>{log.old_value}</span>
                            {" → "}
                            <span style={{color:"#34d399"}}>{log.new_value}</span>
                          </div>
                        )}
                        <div style={{fontSize:11,color:HW.muted,marginTop:4}}>
                          {new Date(log.created_at).toLocaleDateString("en-GB",{
                            day:"numeric",month:"short",year:"numeric",
                            hour:"2-digit",minute:"2-digit"})}
                          {log.logged_by&&` · by ${log.logged_by}`}
                        </div>
                      </div>
                    </div>
                  ))
                }
              </div>
            )}

            {profileTab==="edit"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:18}}>Edit Profile</h3>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                  <div><label style={s.label}>Department</label>
                    <input style={s.input} value={selected.department||""}
                      onChange={e=>setSelected({...selected,department:e.target.value,
                        _original:selected._original||{...selected}})}/></div>
                  <div><label style={s.label}>Assigned Mentor</label>
                    <input style={s.input} value={selected.assigned_mentor||""}
                      onChange={e=>setSelected({...selected,assigned_mentor:e.target.value,
                        _original:selected._original||{...selected}})}/></div>
                  <div><label style={s.label}>GPA</label>
                    <input style={s.input} type="number" step="0.01" min="0" max="4"
                      value={selected.gpa||""}
                      onChange={e=>setSelected({...selected,gpa:e.target.value})}/></div>
                  <div><label style={s.label}>Joining Date</label>
                    <input style={s.input} type="date" value={selected.joining_date||""}
                      onChange={e=>setSelected({...selected,joining_date:e.target.value})}/></div>
                </div>
                <button style={{...s.btn,background:HW.red,color:HW.white,marginTop:16}}
                  onClick={saveProfile}>Save Changes</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                  fontSize:13,marginTop:10}}>{msg}</p>}
              </div>
            )}

            {profileTab==="reports"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>Daily Reports + AI Analysis</h3>
                {reports.length===0?<p style={{color:HW.muted}}>No reports yet.</p>
                  :reports.map(r=>{
                    let pie=null;
                    try{pie=r.pie_chart_json?JSON.parse(r.pie_chart_json):null;}catch(e){}
                    return (
                      <div key={r.id} style={{background:HW.surface2,borderRadius:12,
                        padding:16,marginBottom:14}}>
                        <div style={{display:"flex",justifyContent:"space-between",
                          alignItems:"center",marginBottom:10}}>
                          <div style={{fontWeight:600}}>📅 {r.report_date}</div>
                          <div style={{display:"flex",alignItems:"center",gap:12}}>
                            {r.penalty_applied&&(
                              <span style={{fontSize:11,color:HW.red,
                                background:"rgba(207,10,44,.1)",
                                padding:"2px 8px",borderRadius:10,fontWeight:700}}>
                                ⚠️ -{r.penalty_amount}%
                              </span>
                            )}
                            {r.kpi_score&&(
                              <div style={{fontSize:22,fontWeight:800,color:kpiColor(r.kpi_score)}}>
                                {r.kpi_score}
                                <span style={{fontSize:11,color:HW.muted,fontWeight:400}}> KPI</span>
                              </div>
                            )}
                          </div>
                        </div>
                        {/* Time In / Time Out */}
                        <div style={{display:"flex",gap:16,marginBottom:8,
                          background:HW.surface,borderRadius:8,padding:"8px 12px"}}>
                          <div style={{textAlign:"center"}}>
                            <div style={{fontSize:10,color:HW.muted,fontWeight:700,
                              textTransform:"uppercase",marginBottom:2}}>Time In</div>
                            <div style={{fontSize:16,fontWeight:800,color:"#34d399",
                              fontFamily:"monospace"}}>
                              {r.signin_time||"—"}
                            </div>
                          </div>
                          <div style={{width:1,background:HW.border}}/>
                          <div style={{textAlign:"center"}}>
                            <div style={{fontSize:10,color:HW.muted,fontWeight:700,
                              textTransform:"uppercase",marginBottom:2}}>Time Out</div>
                            <div style={{fontSize:16,fontWeight:800,color:"#4f8ef7",
                              fontFamily:"monospace"}}>
                              {r.signout_time||"—"}
                            </div>
                          </div>
                          <div style={{width:1,background:HW.border}}/>
                          <div style={{textAlign:"center"}}>
                            <div style={{fontSize:10,color:HW.muted,fontWeight:700,
                              textTransform:"uppercase",marginBottom:2}}>Status</div>
                            <div style={{fontSize:13,fontWeight:700,
                              color:r.attended?"#34d399":HW.red}}>
                              {r.attended?"● Present":"○ Absent"}
                            </div>
                          </div>
                        </div>
                        {r.excuse_type&&(
                          <div style={{background:"rgba(255,165,0,.08)",borderRadius:8,
                            padding:10,marginBottom:10,borderLeft:"3px solid #FFA500"}}>
                            <div style={{fontSize:11,color:"#FFA500",fontWeight:700,marginBottom:4}}>
                              Excuse: {r.excuse_type}
                            </div>
                            <div style={{fontSize:12}}>{r.excuse_text}</div>
                          </div>
                        )}
                        <div style={{fontSize:13,marginBottom:12,
                          borderLeft:`3px solid ${HW.red}`,paddingLeft:10}}>
                          {r.report_text}
                        </div>
                        {r.face_capture_url&&(
                          <div style={{marginBottom:10}}>
                            <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                              textTransform:"uppercase",marginBottom:6}}>🤳 Face ID</div>
                            <img src={r.face_capture_url} alt="face"
                              style={{width:80,height:80,objectFit:"cover",
                                borderRadius:"50%",border:`2px solid ${HW.red}`}}/>
                          </div>
                        )}
                        {pie&&(
                          <div style={{marginBottom:12}}>
                            <div style={{fontSize:11,color:HW.muted,fontWeight:700,
                              textTransform:"uppercase",marginBottom:8}}>Task Breakdown</div>
                            <PieChart data={pie}/>
                          </div>
                        )}
                        {r.talent_notes&&(
                          <div style={{background:"rgba(207,10,44,.06)",borderRadius:8,
                            padding:10,borderLeft:`3px solid ${HW.red}`}}>
                            <div style={{fontSize:11,color:HW.red,fontWeight:700,marginBottom:4}}>
                              🌟 AI Talent Notes
                            </div>
                            <div style={{fontSize:12,lineHeight:1.6}}>{r.talent_notes}</div>
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
                <h3 style={{marginBottom:16}}>⚠️ Penalty Log</h3>
                {penalties.length===0?<p style={{color:HW.muted}}>No penalties recorded.</p>:(
                  <>
                    <div style={{background:"rgba(207,10,44,.08)",
                      border:`1px solid rgba(207,10,44,.2)`,
                      borderRadius:10,padding:14,marginBottom:16}}>
                      <div style={{fontSize:13,color:HW.red,fontWeight:700}}>
                        Total: {penalties.length} day(s) ·{" "}
                        {(penalties.length*PENALTY_PCT).toFixed(2)}% deduction
                      </div>
                    </div>
                    {penalties.map(p=>(
                      <div key={p.id} style={{display:"flex",justifyContent:"space-between",
                        alignItems:"center",background:HW.surface2,borderRadius:10,
                        borderLeft:`3px solid ${HW.red}`,padding:12,marginBottom:10}}>
                        <div>
                          <div style={{fontSize:13,fontWeight:600}}>📅 {p.report_date}</div>
                          <div style={{fontSize:12,color:HW.muted,marginTop:3}}>{p.reason}</div>
                        </div>
                        <div style={{fontSize:18,fontWeight:800,color:HW.red}}>-{p.amount}%</div>
                      </div>
                    ))}
                  </>
                )}
              </div>
            )}

            {profileTab==="goals"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>🎯 Trainee Goals</h3>
                <div style={{display:"grid",
                  gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",
                  gap:10,marginBottom:20}}>
                  {KRA_CATEGORIES.map(kra=>{
                    const kraGoals=selectedGoals.filter(g=>g.kra===kra.id);
                    const completed=kraGoals.filter(g=>g.status==="completed").length;
                    return (
                      <div key={kra.id} style={{background:HW.surface2,borderRadius:10,
                        padding:12,borderTop:`3px solid ${kra.color}`,textAlign:"center"}}>
                        <div style={{fontSize:18,marginBottom:4}}>{kra.icon}</div>
                        <div style={{fontSize:10,color:kra.color,fontWeight:700,
                          textTransform:"uppercase",marginBottom:4}}>
                          {kra.label.split(" ")[0]}
                        </div>
                        <div style={{fontSize:13,fontWeight:700}}>{kraGoals.length} goals</div>
                        <div style={{fontSize:11,color:"#34d399"}}>{completed} done</div>
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
          </div>
        )}

        {/* ANALYTICS */}
        {mgmtTab==="analytics"&&(
          <div>
            <div style={{display:"grid",
              gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",
              gap:16,marginBottom:24}}>
              {[
                {label:"Active Trainees",value:analytics.active.length,icon:"👥",color:HW.red},
                {label:"Avg KPI Score",value:analytics.avgKpi,icon:"📊",color:"#FFA500"},
                {label:"Attendance Rate",value:`${analytics.attendanceRate}%`,icon:"✅",color:"#34d399"},
                {label:"Total Penalties",value:analytics.totalPenalties,icon:"⚠️",color:"#f87171"},
                {label:"Total Reports",value:allReports.length,icon:"📋",color:HW.red},
                {label:"Departments",
                  value:[...new Set(trainees.map(t=>t.department).filter(Boolean))].length,
                  icon:"🏢",color:"#FFA500"},
              ].map((stat,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:14,padding:20,borderTop:`3px solid ${stat.color}`}}>
                  <div style={{fontSize:24,marginBottom:8}}>{stat.icon}</div>
                  <div style={{fontSize:28,fontWeight:800,color:stat.color}}>{stat.value}</div>
                  <div style={{fontSize:12,color:HW.muted,marginTop:4}}>{stat.label}</div>
                </div>
              ))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:20}}>
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>🏆 KPI Leaderboard</h3>
                {analytics.traineeKpi.length===0?<p style={{color:HW.muted}}>No data yet.</p>
                  :analytics.traineeKpi.slice(0,8).map((t,i)=>(
                    <div key={t.id} style={{display:"flex",alignItems:"center",gap:12,
                      padding:"10px 0",borderBottom:`1px solid ${HW.border}`}}>
                      <div style={{width:28,height:28,borderRadius:"50%",
                        background:i===0?"#FFD700":i===1?"#C0C0C0":i===2?"#CD7F32":HW.surface2,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        fontSize:12,fontWeight:800,color:i<3?HW.dark:HW.muted,flexShrink:0}}>
                        {i+1}
                      </div>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:600,fontSize:14}}>{t.full_name}</div>
                        <div style={{fontSize:11,color:HW.muted}}>
                          {t.department} · {t.reports} reports
                        </div>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <div style={{fontSize:18,fontWeight:800,color:kpiColor(parseFloat(t.avgKpi))}}>
                          {t.avgKpi}
                        </div>
                        <div style={{fontSize:10,color:HW.muted}}>avg KPI</div>
                      </div>
                    </div>
                  ))
                }
              </div>
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>🏢 Department Performance</h3>
                {analytics.deptData.length===0?<p style={{color:HW.muted}}>No data yet.</p>
                  :<BarChart data={analytics.deptData}/>}
                <div style={{marginTop:16}}>
                  {analytics.deptData.map((d,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",
                      alignItems:"center",padding:"8px 0",
                      borderBottom:`1px solid ${HW.border}`}}>
                      <span style={{fontSize:13,color:HW.muted}}>{d.label}</span>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:80,height:6,background:HW.border,
                          borderRadius:10,overflow:"hidden"}}>
                          <div style={{height:"100%",background:HW.red,
                            borderRadius:10,width:`${d.value}%`}}/>
                        </div>
                        <span style={{fontSize:13,fontWeight:700,
                          color:kpiColor(d.value),minWidth:30}}>{d.value}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            {analytics.atRisk.length>0&&(
              <div style={{...s.card,border:`1px solid rgba(248,113,113,.3)`}}>
                <h3 style={{marginBottom:16,color:"#f87171"}}>⚠️ At Risk Trainees</h3>
                <div style={{display:"grid",
                  gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:12}}>
                  {analytics.atRisk.map(t=>(
                    <div key={t.id} style={{background:HW.surface2,borderRadius:12,
                      padding:16,borderLeft:"3px solid #f87171"}}>
                      <div style={{fontWeight:700,fontSize:15,marginBottom:4}}>{t.full_name}</div>
                      <div style={{fontSize:12,color:HW.muted,marginBottom:8}}>{t.department}</div>
                      <div style={{display:"flex",gap:12}}>
                        <div style={{textAlign:"center"}}>
                          <div style={{fontSize:20,fontWeight:800,color:"#f87171"}}>{t.avgKpi}</div>
                          <div style={{fontSize:10,color:HW.muted}}>Avg KPI</div>
                        </div>
                        <div style={{textAlign:"center"}}>
                          <div style={{fontSize:20,fontWeight:800,color:"#f87171"}}>{t.penalties}</div>
                          <div style={{fontSize:10,color:HW.muted}}>Penalties</div>
                        </div>
                        <div style={{textAlign:"center"}}>
                          <div style={{fontSize:20,fontWeight:800,color:HW.muted}}>{t.reports}</div>
                          <div style={{fontSize:10,color:HW.muted}}>Reports</div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* OKR */}
        {mgmtTab==="okr"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",marginBottom:20}}>
              <div>
                <h3 style={{margin:0}}>🎯 OKR Tracking</h3>
                <p style={{color:HW.muted,fontSize:13,margin:"4px 0 0"}}>
                  Objectives and Key Results by Department
                </p>
              </div>
              <button style={{...s.btn,background:HW.red,color:HW.white}}
                onClick={()=>setShowAddOkr(!showAddOkr)}>
                {showAddOkr?"Cancel":"+ Add OKR"}
              </button>
            </div>
            {showAddOkr&&(
              <div style={{...s.card,border:`1px solid rgba(207,10,44,.3)`}}>
                <h4 style={{marginBottom:16,color:HW.red}}>Add New OKR</h4>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                  <div><label style={s.label}>Department</label>
                    <input style={s.input} placeholder="e.g. Engineering"
                      value={newOkr.department}
                      onChange={e=>setNewOkr({...newOkr,department:e.target.value})}/></div>
                  <div><label style={s.label}>Due Date</label>
                    <input style={s.input} type="date" value={newOkr.due_date}
                      onChange={e=>setNewOkr({...newOkr,due_date:e.target.value})}/></div>
                  <div style={{gridColumn:"1/-1"}}><label style={s.label}>Objective</label>
                    <input style={s.input} placeholder="e.g. Improve Technical Skills"
                      value={newOkr.objective}
                      onChange={e=>setNewOkr({...newOkr,objective:e.target.value})}/></div>
                  <div style={{gridColumn:"1/-1"}}><label style={s.label}>Key Result</label>
                    <input style={s.input} placeholder="e.g. Complete 20 lab sessions"
                      value={newOkr.key_result}
                      onChange={e=>setNewOkr({...newOkr,key_result:e.target.value})}/></div>
                  <div><label style={s.label}>Target</label>
                    <input style={s.input} type="number" value={newOkr.target}
                      onChange={e=>setNewOkr({...newOkr,target:parseFloat(e.target.value)})}/></div>
                  <div><label style={s.label}>Unit</label>
                    <select style={s.input} value={newOkr.unit}
                      onChange={e=>setNewOkr({...newOkr,unit:e.target.value})}>
                      <option value="%">%</option>
                      <option value="sessions">Sessions</option>
                      <option value="reports">Reports</option>
                      <option value="tasks">Tasks</option>
                      <option value="reviews">Reviews</option>
                      <option value="workflows">Workflows</option>
                    </select></div>
                </div>
                <button style={{...s.btn,background:HW.red,color:HW.white,marginTop:16}}
                  onClick={addOkr}>Add OKR</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
                  fontSize:13,marginTop:10}}>{msg}</p>}
              </div>
            )}
            <div style={{display:"grid",
              gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",
              gap:16,marginBottom:20}}>
              {[
                {label:"Total OKRs",value:okrs.length,color:HW.red},
                {label:"On Track",value:okrs.filter(o=>(o.current/o.target)>=.8).length,color:"#34d399"},
                {label:"In Progress",value:okrs.filter(o=>(o.current/o.target)>=.5&&(o.current/o.target)<.8).length,color:"#FFA500"},
                {label:"Behind",value:okrs.filter(o=>(o.current/o.target)<.5).length,color:"#f87171"},
              ].map((s2,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:14,padding:16,textAlign:"center",
                  borderTop:`3px solid ${s2.color}`}}>
                  <div style={{fontSize:28,fontWeight:800,color:s2.color}}>{s2.value}</div>
                  <div style={{fontSize:12,color:HW.muted,marginTop:4}}>{s2.label}</div>
                </div>
              ))}
            </div>
            {[...new Set(okrs.map(o=>o.department))].map(dept=>(
              <div key={dept} style={s.card}>
                <h4 style={{marginBottom:16,color:HW.red}}>🏢 {dept}</h4>
                {okrs.filter(o=>o.department===dept).map(okr=>(
                  <OKRBar key={okr.id} okr={okr} onUpdate={updateOkr}/>
                ))}
              </div>
            ))}
            {okrs.length===0&&(
              <div style={{...s.card,textAlign:"center",padding:40}}>
                <div style={{fontSize:40,marginBottom:12}}>🎯</div>
                <p style={{color:HW.muted}}>No OKRs yet. Click "+ Add OKR"!</p>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }
}