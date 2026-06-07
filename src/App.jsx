import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";

const CLAUDE_KEY  = import.meta.env.VITE_CLAUDE_API_KEY;
const WORK_LAT    = 23.5957;
const WORK_LNG    = 58.3911;
const WORK_RADIUS = 500;
const MAX_SIGNIN  = "09:00";
const PENALTY_PCT = 8.33;

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
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLng/2)**2;
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

// ── Bar Chart ────────────────────────────────────────
function BarChart({data,color=HW.red}) {
  const max=Math.max(...data.map(d=>d.value),1);
  return (
    <div style={{display:"flex",alignItems:"flex-end",gap:8,height:120,padding:"0 4px"}}>
      {data.map((d,i)=>(
        <div key={i} style={{flex:1,display:"flex",flexDirection:"column",
          alignItems:"center",gap:4}}>
          <div style={{fontSize:10,color:HW.muted,fontWeight:700}}>{d.value}</div>
          <div style={{width:"100%",background:color,borderRadius:"4px 4px 0 0",
            height:`${(d.value/max)*80}px`,minHeight:4,transition:"height .5s ease"}}/>
          <div style={{fontSize:9,color:HW.muted,textAlign:"center",
            whiteSpace:"nowrap",overflow:"hidden",maxWidth:40}}>{d.label}</div>
        </div>
      ))}
    </div>
  );
}

// ── OKR Progress Bar ─────────────────────────────────
function OKRBar({okr,onUpdate,isManager}) {
  const pct=Math.min((okr.current/okr.target)*100,100).toFixed(0);
  const color=pct>=80?HW.red:pct>=50?"#FFA500":"#666";
  return (
    <div style={{background:HW.surface2,borderRadius:12,padding:16,marginBottom:12}}>
      <div style={{display:"flex",justifyContent:"space-between",
        alignItems:"flex-start",marginBottom:8}}>
        <div style={{flex:1}}>
          <div style={{fontSize:12,color:HW.muted,marginBottom:2}}>{okr.department}</div>
          <div style={{fontWeight:600,fontSize:14,marginBottom:2}}>{okr.objective}</div>
          <div style={{fontSize:13,color:HW.muted}}>{okr.key_result}</div>
        </div>
        <div style={{textAlign:"right",marginLeft:12}}>
          <div style={{fontSize:22,fontWeight:800,color}}>{pct}%</div>
          <div style={{fontSize:11,color:HW.muted}}>
            {okr.current}/{okr.target} {okr.unit}
          </div>
        </div>
      </div>
      <div style={{height:8,background:HW.border,borderRadius:10,overflow:"hidden"}}>
        <div style={{height:"100%",borderRadius:10,background:color,
          width:`${pct}%`,transition:"width .6s ease"}}/>
      </div>
      {okr.due_date&&(
        <div style={{fontSize:11,color:HW.muted,marginTop:6}}>
          Due: {new Date(okr.due_date).toLocaleDateString("en-GB",{
            day:"numeric",month:"short",year:"numeric"})}
        </div>
      )}
      {isManager&&onUpdate&&(
        <div style={{display:"flex",gap:8,marginTop:10,alignItems:"center"}}>
          <input type="number" defaultValue={okr.current}
            style={{background:HW.surface,border:`1px solid ${HW.border}`,
              color:HW.text,borderRadius:6,padding:"4px 8px",
              width:80,fontSize:13}}
            onBlur={e=>onUpdate(okr.id,parseFloat(e.target.value))}/>
          <span style={{fontSize:12,color:HW.muted}}>Update progress</span>
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
        <button onClick={onDismiss} style={{background:HW.red,color:HW.white,
          border:"none",borderRadius:8,padding:"12px 32px",fontWeight:700,
          fontSize:14,cursor:"pointer",width:"100%"}}>
          Got it — I'll submit now!
        </button>
      </div>
    </div>
  );
}

export default function App() {
  const [view,setView]               = useState("login");
  const [mgmtTab,setMgmtTab]         = useState("trainees"); // trainees | analytics
  const [user,setUser]               = useState(null);
  const [trainees,setTrainees]       = useState([]);
  const [selected,setSelected]       = useState(null);
  const [reports,setReports]         = useState([]);
  const [logs,setLogs]               = useState([]);
  const [penalties,setPenalties]     = useState([]);
  const [okrs,setOkrs]               = useState([]);
  const [allReports,setAllReports]   = useState([]);
  const [allPenalties,setAllPenalties] = useState([]);
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
  const [trafficCount,setTrafficCount] = useState(0);
  const [faceCapture,setFaceCapture] = useState(null);
  const [faceCaptured,setFaceCaptured] = useState(false);
  const [cameraActive,setCameraActive] = useState(false);
  const [isLate,setIsLate]           = useState(false);
  const [excuseType,setExcuseType]   = useState("");
  const [excuseText,setExcuseText]   = useState("");
  const [excusePhoto,setExcusePhoto] = useState(null);
  const [excusePreview,setExcusePreview] = useState(null);
  const [newOkr,setNewOkr]           = useState({department:"",objective:"",
    key_result:"",target:100,current:0,unit:"%",due_date:""});
  const [showAddOkr,setShowAddOkr]   = useState(false);
  const videoRef=useRef(),canvasRef=useRef(),streamRef=useRef();
  const excuseRef=useRef(),photoRef=useRef();

  const [profile,setProfile] = useState({
    full_name:"",civil_id:"",phone_number:"",
    department:"",assigned_mentor:"",gpa:"",
  });
  const [report,setReport] = useState({
    report_date:new Date().toISOString().split("T")[0],
    attended:false,signin_time:"",report_text:""
  });

  const s = {
    page:{minHeight:"100vh",background:HW.dark,color:HW.text,fontFamily:"sans-serif",padding:32},
    card:{background:HW.surface,border:`1px solid ${HW.border}`,borderRadius:16,padding:24,marginBottom:20},
    input:{background:HW.surface2,border:`1px solid ${HW.border}`,color:HW.text,borderRadius:8,
      padding:"9px 13px",width:"100%",fontFamily:"inherit",fontSize:14,boxSizing:"border-box"},
    label:{fontSize:11,fontWeight:700,color:HW.muted,textTransform:"uppercase",
      letterSpacing:".06em",display:"block",marginBottom:5},
    btn:{padding:"10px 20px",borderRadius:8,border:"none",fontWeight:700,cursor:"pointer",fontSize:13},
    th:{textAlign:"left",padding:"10px 14px",fontSize:11,fontWeight:700,color:HW.muted,
      textTransform:"uppercase",borderBottom:`1px solid ${HW.border}`},
    td:{padding:"12px 14px",fontSize:14,borderBottom:`1px solid rgba(51,51,51,.6)`},
  };

  const statusColors = {
    active:{bg:"rgba(207,10,44,.15)",color:HW.red},
    inactive:{bg:"rgba(136,136,136,.15)",color:HW.muted},
    transferred:{bg:"rgba(255,165,0,.15)",color:"#FFA500"},
    dropped:{bg:"rgba(136,136,136,.2)",color:"#666"},
  };

  const eventIcons = {
    joined:"🟢",dropped:"🔴",transferred:"🔄",
    mentor_changed:"👤",dept_changed:"🏢",
    reactivated:"✅",note:"📝",inactive:"⏸"
  };

  function kpiColor(score) {
    return score>=80?HW.red:score>=60?"#FFA500":"#666";
  }

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

  async function handleSession(authUser) {
    setUser(authUser);
    const{data}=await supabase.from("profiles")
      .select("role,trainee_id").eq("id",authUser.id).single();
    if(data){
      if(data.role==="management"){
        setView("mgmt");fetchTrainees();fetchOkrs();fetchAllData();
      } else {
        setView("trainee");
        if(data.trainee_id){
          const{data:t}=await supabase.from("trainees")
            .select("*").eq("id",data.trainee_id).single();
          if(t) setProfile({full_name:t.full_name||"",civil_id:t.civil_id||"",
            phone_number:t.phone_number||"",department:t.department||"",
            assigned_mentor:t.assigned_mentor||"",gpa:t.gpa||""});
          const start=new Date();start.setDate(1);
          const{data:tc}=await supabase.from("traffic_excuses")
            .select("id").eq("trainee_id",data.trainee_id)
            .gte("created_at",start.toISOString());
          if(tc) setTrafficCount(tc.length);
        }
      }
    }
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
    await supabase.from("profiles").insert({id:authData.user.id,email,role:"trainee"});
    setMsg("✅ Account created! You can now log in.");
    setLoading(false);
    setTimeout(()=>{setView("login");setMsg("");setConfirmPwd("");},2000);
  }

  async function logout(){
    await supabase.auth.signOut();
    setView("login");setUser(null);setSelected(null);setMsg("");
    setAiResult(null);setPhotoFile(null);setPhotoPreview(null);
    setLocationOk(false);setGeoStatus("idle");setGeoMsg("");
    setFaceCapture(null);setFaceCaptured(false);stopCamera();
  }

  function checkLocation(){
    setGeoStatus("checking");setGeoMsg("📍 Checking your location…");
    if(!navigator.geolocation){setGeoStatus("error");setGeoMsg("❌ Browser does not support location.");return;}
    navigator.geolocation.getCurrentPosition(
      (pos)=>{
        const dist=getDistance(pos.coords.latitude,pos.coords.longitude,WORK_LAT,WORK_LNG);
        if(dist<=WORK_RADIUS){setLocationOk(true);setGeoStatus("ok");
          setGeoMsg(`✅ Verified — ${Math.round(dist)}m from workplace.`);}
        else{setLocationOk(false);setGeoStatus("error");
          setGeoMsg(`❌ You are ${Math.round(dist)}m away. Must be within 500m.`);}
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
    canvas.toBlob(blob=>{setFaceCapture(blob);setFaceCaptured(true);stopCamera();
      setMsg("✅ Face captured!");},"image/jpeg",0.8);
  }

  function handleProofPhoto(e){const f=e.target.files[0];if(!f)return;setPhotoFile(f);setPhotoPreview(URL.createObjectURL(f));}
  function handleExcusePhoto(e){const f=e.target.files[0];if(!f)return;setExcusePhoto(f);setExcusePreview(URL.createObjectURL(f));}

  async function uploadFile(bucket,path,file){
    const{error}=await supabase.storage.from(bucket).upload(path,file);
    if(error) return null;
    return supabase.storage.from(bucket).getPublicUrl(path).data.publicUrl;
  }

  async function submitReport(){
    if(!locationOk){setMsg("📍 Please verify your location first.");return;}
    if(!faceCaptured){setMsg("🤳 Please complete Face ID verification first.");return;}
    if(!profile.full_name||!profile.civil_id){setMsg("Please enter your Full Name and Civil ID.");return;}
    if(!report.report_text){setMsg("Please write your daily report.");return;}
    if(excuseType==="traffic"){
      if(trafficCount>=2){setMsg("❌ Traffic excuse limit reached (max 2 per month).");return;}
      if(!excusePhoto){setMsg("⚠️ Traffic excuse requires a proof photo.");return;}
    }
    if(isLate&&!excuseText){setMsg("⚠️ Please describe your reason for being late.");return;}
    const penaltyApplied=!report.attended||(report.signin_time&&report.signin_time>MAX_SIGNIN&&!isLate);
    setLoading(true);setAiLoading(true);setMsg("");setAiResult(null);

    const{data:traineeData,error:tErr}=await supabase.from("trainees").upsert({
      full_name:profile.full_name,civil_id:profile.civil_id,
      phone_number:profile.phone_number,department:profile.department,
      assigned_mentor:profile.assigned_mentor,
      gpa:profile.gpa?parseFloat(profile.gpa):null,
      joining_date:new Date().toISOString().split("T")[0],
    },{onConflict:"civil_id"}).select().single();
    if(tErr){setMsg("Error: "+tErr.message);setLoading(false);setAiLoading(false);return;}
    await supabase.from("profiles").update({trainee_id:traineeData.id}).eq("id",user.id);

    const tid=traineeData.id,ts=Date.now();
    setMsg("🤳 Uploading face capture…");
    const faceUrl=faceCapture?await uploadFile("report-photos",`${tid}/face_${ts}.jpg`,faceCapture):null;
    setMsg("📸 Uploading proof photo…");
    const photoUrl=photoFile?await uploadFile("report-photos",`${tid}/proof_${ts}.jpg`,photoFile):null;
    let excusePhotoUrl=null;
    if(excusePhoto) excusePhotoUrl=await uploadFile("report-photos",`${tid}/excuse_${ts}.jpg`,excusePhoto);
    if(excuseType==="traffic"&&excusePhotoUrl){
      await supabase.from("traffic_excuses").insert({trainee_id:tid,report_date:report.report_date,photo_url:excusePhotoUrl});
      setTrafficCount(c=>c+1);
    }
    setMsg("🤖 AI is analyzing your report…");
    let ai=null;
    try{ai=await analyzeReport(report.report_text);setAiResult(ai);}catch(e){console.error(e);}
    if(penaltyApplied){
      await supabase.from("penalties").insert({
        trainee_id:tid,report_date:report.report_date,
        reason:!report.attended?"Absent — no attendance recorded":"Late sign-in after 9:00 AM without excuse",
        amount:PENALTY_PCT,
      });
    }
    const{error:rErr}=await supabase.from("daily_reports").upsert({
      trainee_id:tid,report_date:report.report_date,attended:report.attended,
      signin_time:report.signin_time||null,report_text:report.report_text,
      photo_url:photoUrl,face_capture_url:faceUrl,
      excuse_type:excuseType||null,excuse_text:excuseText||null,
      excuse_photo_url:excusePhotoUrl,traffic_excuse:excuseType==="traffic",
      kpi_score:ai?.kpi_score||null,
      pie_chart_json:ai?.pie_chart?JSON.stringify(ai.pie_chart):null,
      talent_notes:ai?.talent_notes||null,
      penalty_applied:penaltyApplied,penalty_amount:penaltyApplied?PENALTY_PCT:0,
    },{onConflict:"trainee_id,report_date"});
    if(rErr) setMsg("Report error: "+rErr.message);
    else setMsg(penaltyApplied
      ?`✅ Report submitted — ⚠️ Penalty of ${PENALTY_PCT}% applied.`
      :"✅ Report submitted and AI analysis complete!");
    setLoading(false);setAiLoading(false);
  }

  async function fetchTrainees(){
    setLoading(true);
    const{data}=await supabase.from("trainees").select("*").order("full_name");
    if(data) setTrainees(data);
    setLoading(false);
  }

  async function fetchOkrs(){
    const{data}=await supabase.from("okrs").select("*").order("department");
    if(data) setOkrs(data);
  }

  async function fetchAllData(){
    const{data:reps}=await supabase.from("daily_reports").select("*");
    const{data:pens}=await supabase.from("penalties").select("*");
    if(reps) setAllReports(reps);
    if(pens) setAllPenalties(pens);
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
    await fetchReports(t.id);await fetchLogs(t.id);await fetchPenalties(t.id);
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
    await supabase.from("okrs").update({current:newValue}).eq("id",id);
    fetchOkrs();
  }

  async function addOkr(){
    if(!newOkr.department||!newOkr.objective||!newOkr.key_result){
      setMsg("Please fill all OKR fields.");return;}
    await supabase.from("okrs").insert({...newOkr,created_by:user?.email});
    setShowAddOkr(false);
    setNewOkr({department:"",objective:"",key_result:"",target:100,current:0,unit:"%",due_date:""});
    fetchOkrs();setMsg("✅ OKR added successfully!");
  }

  async function exportPDF(trainee){
    const doc=new jsPDF();
    const t=trainee||selected;
    doc.setFillColor(207,10,44);
    doc.rect(0,0,210,40,"F");
    doc.setTextColor(255,255,255);
    doc.setFontSize(20);doc.setFont("helvetica","bold");
    doc.text("TraineeOS — Weekly Performance Report",14,18);
    doc.setFontSize(11);doc.setFont("helvetica","normal");
    doc.text(`Generated: ${new Date().toLocaleDateString("en-GB")}`,14,30);
    doc.setTextColor(0,0,0);
    doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text("Trainee Information",14,52);
    autoTable(doc,{
      startY:56,
      head:[["Field","Details"]],
      body:[
        ["Full Name",t.full_name||"—"],["Civil ID",t.civil_id||"—"],
        ["Department",t.department||"—"],["Assigned Mentor",t.assigned_mentor||"—"],
        ["GPA",t.gpa?.toString()||"—"],["Status",t.status||"—"],
        ["Joining Date",t.joining_date||"—"],
        ["Quitting Date",t.quitting_date||"Still Active"],
        ["Days in Program",t.joining_date
          ?Math.floor((new Date()-new Date(t.joining_date))/(1000*60*60*24))+" days":"—"],
      ],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    const reportsY=doc.lastAutoTable.finalY+10;
    doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text("Weekly Reports",14,reportsY);
    const{data:reps}=await supabase.from("daily_reports").select("*")
      .eq("trainee_id",t.id).order("report_date",{ascending:false}).limit(7);
    if(reps&&reps.length>0){
      autoTable(doc,{
        startY:reportsY+4,
        head:[["Date","Attended","Sign-in","KPI","Penalty","Report"]],
        body:reps.map(r=>[r.report_date,r.attended?"✓":"✗",r.signin_time||"—",
          r.kpi_score||"—",r.penalty_applied?`-${r.penalty_amount}%`:"None",
          (r.report_text||"").substring(0,50)+"..."]),
        headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
        alternateRowStyles:{fillColor:[245,245,245]},
      });
    }
    const penY=doc.lastAutoTable?.finalY+10||reportsY+60;
    doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text("Penalty Summary",14,penY);
    const{data:pens}=await supabase.from("penalties").select("*").eq("trainee_id",t.id);
    autoTable(doc,{
      startY:penY+4,
      head:[["Date","Reason","Deduction"]],
      body:pens&&pens.length>0?pens.map(p=>[p.report_date,p.reason,`-${p.amount}%`]):[["—","No penalties","—"]],
      headStyles:{fillColor:[207,10,44],textColor:[255,255,255]},
      alternateRowStyles:{fillColor:[245,245,245]},
    });
    const total=pens?(pens.length*PENALTY_PCT).toFixed(2):0;
    doc.setTextColor(207,10,44);doc.setFontSize(11);doc.setFont("helvetica","bold");
    doc.text(`Total Deduction: ${total}%`,14,doc.lastAutoTable.finalY+6);
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
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allT||[]).map(t=>({
      "Full Name":t.full_name,"Civil ID":t.civil_id,"Phone":t.phone_number,
      "Department":t.department,"Mentor":t.assigned_mentor,"GPA":t.gpa,
      "Status":t.status,"Joining Date":t.joining_date,
      "Quitting Date":t.quitting_date||"Still Active",
      "Days in Program":t.joining_date?Math.floor((new Date()-new Date(t.joining_date))/(1000*60*60*24)):"—",
    }))),"Trainees");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allR||[]).map(r=>({
      "Trainee ID":r.trainee_id,"Date":r.report_date,"Attended":r.attended?"Yes":"No",
      "Sign-in":r.signin_time||"—","KPI Score":r.kpi_score||"—",
      "Report":r.report_text,"Talent Notes":r.talent_notes||"—",
      "Penalty":r.penalty_applied?`-${r.penalty_amount}%`:"None",
      "Excuse":r.excuse_type||"—",
    }))),"Daily Reports");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allP||[]).map(p=>({
      "Trainee ID":p.trainee_id,"Date":p.report_date,"Reason":p.reason,
      "Deduction":`-${p.amount}%`,
    }))),"Penalties");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allL||[]).map(l=>({
      "Trainee ID":l.trainee_id,"Event":l.event_type,"Description":l.description,
      "Old":l.old_value||"—","New":l.new_value||"—","By":l.logged_by,
      "Date":new Date(l.created_at).toLocaleDateString("en-GB"),
    }))),"Timeline");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet((allT||[]).map(t=>{
      const tr=(allR||[]).filter(r=>r.trainee_id===t.id);
      const tp=(allP||[]).filter(p=>p.trainee_id===t.id);
      const avg=tr.filter(r=>r.kpi_score).reduce((a,r,_,arr)=>a+r.kpi_score/arr.length,0);
      return{"Full Name":t.full_name,"Department":t.department,"Status":t.status,
        "Joining":t.joining_date||"—","Quitting":t.quitting_date||"Active",
        "Reports":tr.length,"Attended":tr.filter(r=>r.attended).length,
        "Absent":tr.filter(r=>!r.attended).length,
        "Avg KPI":avg?avg.toFixed(1):"—","Penalties":tp.length,
        "Deduction":`${(tp.length*PENALTY_PCT).toFixed(2)}%`};
    })),"Summary");
    XLSX.writeFile(wb,`TraineeOS_Export_${new Date().toISOString().split("T")[0]}.xlsx`);
    setMsg("✅ Excel exported successfully!");
  }

  // ── Analytics calculations ───────────────────────────
  function getAnalytics() {
    const active=trainees.filter(t=>t.status==="active");
    const totalReports=allReports.length;
    const attended=allReports.filter(r=>r.attended).length;
    const attendanceRate=totalReports>0?((attended/totalReports)*100).toFixed(1):0;
    const kpiReports=allReports.filter(r=>r.kpi_score);
    const avgKpi=kpiReports.length>0
      ?(kpiReports.reduce((a,r)=>a+r.kpi_score,0)/kpiReports.length).toFixed(1):0;
    const totalPenalties=allPenalties.length;

    // KPI per trainee for leaderboard
    const traineeKpi=trainees.map(t=>{
      const tr=allReports.filter(r=>r.trainee_id===t.id&&r.kpi_score);
      const avg=tr.length>0?tr.reduce((a,r)=>a+r.kpi_score,0)/tr.length:0;
      const tp=allPenalties.filter(p=>p.trainee_id===t.id).length;
      return{...t,avgKpi:avg.toFixed(1),penalties:tp,reports:allReports.filter(r=>r.trainee_id===t.id).length};
    }).sort((a,b)=>b.avgKpi-a.avgKpi);

    // Dept performance
    const depts=[...new Set(trainees.map(t=>t.department).filter(Boolean))];
    const deptData=depts.map(dept=>{
      const dTrainees=trainees.filter(t=>t.department===dept);
      const dReports=allReports.filter(r=>dTrainees.some(t=>t.id===r.trainee_id)&&r.kpi_score);
      const avg=dReports.length>0?dReports.reduce((a,r)=>a+r.kpi_score,0)/dReports.length:0;
      return{label:dept.substring(0,8),value:Math.round(avg)};
    });

    // At risk trainees (KPI < 60 or penalties > 2)
    const atRisk=traineeKpi.filter(t=>parseFloat(t.avgKpi)<60||t.penalties>2);

    return{active,attendanceRate,avgKpi,totalPenalties,traineeKpi,deptData,atRisk};
  }

  // ══════════════════════════════════════════════════
  // LOGIN
  // ══════════════════════════════════════════════════
  if(view==="login") return (
    <div style={{...s.page,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{...s.card,width:420}}>
        <div style={{textAlign:"center",marginBottom:28,paddingBottom:24,borderBottom:`1px solid ${HW.border}`}}>
          <HuaweiLogo size={56}/>
          <h2 style={{margin:"12px 0 4px",fontSize:24}}>TraineeOS</h2>
          <p style={{color:HW.red,fontSize:12,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase"}}>
            Powered by Huawei
          </p>
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
          onClick={()=>{setView("login");setMsg("");}}>
          ← Back to Login
        </button>
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
          fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════
  // TRAINEE DASHBOARD
  // ══════════════════════════════════════════════════
  if(view==="trainee") return (
    <div style={s.page}>
      {showReminder&&<ReminderPopup onDismiss={()=>setShowReminder(false)}/>}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",
        marginBottom:28,paddingBottom:20,borderBottom:`1px solid ${HW.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <HuaweiLogo size={36}/>
          <div>
            <h2 style={{margin:0,fontSize:20}}>Trainee Dashboard</h2>
            <p style={{color:HW.muted,fontSize:12,margin:0}}>{user?.email}</p>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button style={{...s.btn,background:HW.surface2,color:HW.muted,fontSize:12}}
            onClick={()=>setShowReminder(true)}>⏰ Test Reminder</button>
          <button style={{...s.btn,background:HW.red,color:HW.white}} onClick={logout}>Sign out</button>
        </div>
      </div>

      {/* Location */}
      <div style={{...s.card,border:locationOk?`1px solid rgba(207,10,44,.5)`:`1px solid ${HW.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
          <span style={{fontSize:20}}>📍</span>
          <h3 style={{margin:0,fontSize:16}}>Step 1 — Location Verification</h3>
        </div>
        <p style={{color:HW.muted,fontSize:13,marginBottom:14}}>
          Must be within <b style={{color:HW.text}}>500m</b> of workplace.
          Max sign-in: <b style={{color:HW.red}}>9:00 AM</b>.
        </p>
        <button style={{...s.btn,background:locationOk?"rgba(207,10,44,.15)":HW.red,
          color:locationOk?HW.red:HW.white,opacity:geoStatus==="checking"?0.6:1}}
          onClick={checkLocation} disabled={geoStatus==="checking"}>
          {geoStatus==="checking"?"Checking…":locationOk?"✅ Location Verified":"📍 Verify My Location"}
        </button>
        {geoMsg&&<p style={{fontSize:13,marginTop:10,color:geoStatus==="ok"?"#34d399":HW.red}}>{geoMsg}</p>}
      </div>

      {/* Face ID */}
      <div style={{...s.card,border:faceCaptured?`1px solid rgba(207,10,44,.5)`:`1px solid ${HW.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
          <span style={{fontSize:20}}>🤳</span>
          <h3 style={{margin:0,fontSize:16}}>Step 2 — Face ID Verification</h3>
        </div>
        {!cameraActive&&!faceCaptured&&(
          <button style={{...s.btn,background:HW.red,color:HW.white}} onClick={startCamera}>📷 Open Camera</button>
        )}
        {cameraActive&&(
          <div>
            <video ref={videoRef} autoPlay playsInline
              style={{width:"100%",maxWidth:320,borderRadius:12,border:`2px solid ${HW.red}`,display:"block",marginBottom:12}}/>
            <canvas ref={canvasRef} style={{display:"none"}}/>
            <div style={{display:"flex",gap:8}}>
              <button style={{...s.btn,background:HW.red,color:HW.white}} onClick={captureface}>📸 Capture Face</button>
              <button style={{...s.btn,background:HW.surface2,color:HW.muted,border:`1px solid ${HW.border}`}} onClick={stopCamera}>Cancel</button>
            </div>
          </div>
        )}
        {faceCaptured&&(
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:60,height:60,borderRadius:"50%",background:"rgba(207,10,44,.15)",
              display:"flex",alignItems:"center",justifyContent:"center",fontSize:28}}>✅</div>
            <div>
              <div style={{fontWeight:700,color:HW.red}}>Face Captured Successfully</div>
              <button style={{...s.btn,background:"none",color:HW.muted,padding:"4px 0",fontSize:12}}
                onClick={()=>{setFaceCaptured(false);setFaceCapture(null);}}>Retake</button>
            </div>
          </div>
        )}
      </div>

      {/* Profile */}
      <div style={s.card}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:18}}>
          <span style={{fontSize:20}}>👤</span>
          <h3 style={{margin:0,fontSize:16}}>Step 3 — My Profile</h3>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div><label style={s.label}>Full Name *</label>
            <input style={s.input} placeholder="Ahmed Al-Rashidi" value={profile.full_name}
              onChange={e=>setProfile({...profile,full_name:e.target.value})}/></div>
          <div><label style={s.label}>Civil ID *</label>
            <input style={s.input} placeholder="10234567" value={profile.civil_id}
              onChange={e=>setProfile({...profile,civil_id:e.target.value})}/></div>
          <div><label style={s.label}>Phone Number</label>
            <input style={s.input} placeholder="+968-9100-0001" value={profile.phone_number}
              onChange={e=>setProfile({...profile,phone_number:e.target.value})}/></div>
          <div><label style={s.label}>Department</label>
            <input style={s.input} placeholder="e.g. Engineering" value={profile.department}
              onChange={e=>setProfile({...profile,department:e.target.value})}/></div>
          <div><label style={s.label}>Assigned Mentor</label>
            <input style={s.input} placeholder="Dr. Fatima Al-Sayed" value={profile.assigned_mentor}
              onChange={e=>setProfile({...profile,assigned_mentor:e.target.value})}/></div>
          <div><label style={s.label}>GPA</label>
            <input style={s.input} type="number" placeholder="3.85" min="0" max="4" step="0.01"
              value={profile.gpa} onChange={e=>setProfile({...profile,gpa:e.target.value})}/></div>
        </div>
      </div>

      {/* Report */}
      <div style={s.card}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:18}}>
          <span style={{fontSize:20}}>📝</span>
          <h3 style={{margin:0,fontSize:16}}>Step 4 — Daily Report</h3>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div><label style={s.label}>Report Date</label>
            <input style={s.input} type="date" value={report.report_date}
              onChange={e=>setReport({...report,report_date:e.target.value})}/></div>
          <div><label style={s.label}>Sign-in Time</label>
            <input style={s.input} type="time" value={report.signin_time}
              onChange={e=>{setReport({...report,signin_time:e.target.value});setIsLate(e.target.value>MAX_SIGNIN);}}/></div>
          <div style={{display:"flex",alignItems:"center",gap:10,paddingTop:22}}>
            <input type="checkbox" id="att" checked={report.attended}
              onChange={e=>setReport({...report,attended:e.target.checked})}/>
            <label htmlFor="att" style={{...s.label,margin:0}}>Attended today</label>
          </div>
          {(!report.attended||(report.signin_time&&report.signin_time>MAX_SIGNIN&&!isLate))&&(
            <div style={{background:"rgba(207,10,44,.1)",border:`1px solid rgba(207,10,44,.3)`,
              borderRadius:8,padding:10,fontSize:12,color:HW.red}}>
              ⚠️ <b>Penalty Warning:</b> {!report.attended?"Not marked as attended":"Sign-in after 9:00 AM"} — <b>8.33% deduction</b>
            </div>
          )}
          <div style={{gridColumn:"1/-1"}}>
            <label style={s.label}>What did you work on today? *</label>
            <textarea style={{...s.input,height:120,resize:"vertical"}}
              placeholder="Describe your tasks, achievements and challenges…"
              value={report.report_text}
              onChange={e=>setReport({...report,report_text:e.target.value})}/>
          </div>
          {isLate&&(
            <div style={{gridColumn:"1/-1",background:HW.surface2,borderRadius:12,padding:16,
              border:`1px solid rgba(207,10,44,.3)`}}>
              <div style={{fontWeight:700,color:HW.red,marginBottom:12}}>⚠️ Late Arrival — Please Provide Excuse</div>
              <div style={{marginBottom:10}}>
                <label style={s.label}>Excuse Type</label>
                <select style={s.input} value={excuseType} onChange={e=>setExcuseType(e.target.value)}>
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
                  placeholder="Please describe your reason for being late…"
                  value={excuseText} onChange={e=>setExcuseText(e.target.value)}/>
              </div>
              <div>
                <label style={s.label}>Proof Photo {excuseType==="traffic"?"(Required)":"(Optional)"}</label>
                <div style={{display:"flex",gap:10,alignItems:"center"}}>
                  <button style={{...s.btn,background:HW.surface,color:HW.text,border:`1px dashed ${HW.border}`}}
                    onClick={()=>excuseRef.current.click()}>
                    📷 {excusePhoto?"Change":"Upload Proof"}
                  </button>
                  <input ref={excuseRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleExcusePhoto}/>
                  {excusePreview&&<img src={excusePreview} alt="proof" style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>}
                </div>
                {excuseType==="traffic"&&(
                  <p style={{fontSize:11,color:HW.muted,marginTop:6}}>Traffic excuses: <b style={{color:trafficCount>=2?HW.red:HW.text}}>{trafficCount}/2</b></p>
                )}
              </div>
            </div>
          )}
          <div style={{gridColumn:"1/-1"}}>
            <label style={s.label}>📸 Proof Photo (optional)</label>
            <div style={{display:"flex",gap:12,alignItems:"center"}}>
              <button style={{...s.btn,background:HW.surface2,color:HW.text,border:`1px dashed ${HW.border}`}}
                onClick={()=>photoRef.current.click()}>
                {photoFile?"📷 Change Photo":"📷 Upload Photo"}
              </button>
              <input ref={photoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleProofPhoto}/>
              {photoPreview&&<img src={photoPreview} alt="preview" style={{width:60,height:60,objectFit:"cover",borderRadius:8}}/>}
            </div>
          </div>
        </div>
        <button style={{...s.btn,
          background:(locationOk&&faceCaptured)?HW.red:HW.surface2,
          color:(locationOk&&faceCaptured)?HW.white:HW.muted,
          marginTop:16,width:"100%",padding:14,
          opacity:(loading||aiLoading)?0.6:1,
          cursor:(locationOk&&faceCaptured)?"pointer":"not-allowed",fontSize:15}}
          onClick={submitReport} disabled={loading||aiLoading||!locationOk||!faceCaptured}>
          {aiLoading?"🤖 AI Analyzing…":loading?"Saving…":"Submit Report ✓"}
        </button>
        {(!locationOk||!faceCaptured)&&(
          <p style={{fontSize:12,color:HW.muted,textAlign:"center",marginTop:8}}>
            {!locationOk&&!faceCaptured?"Complete location + face verification above"
              :!locationOk?"📍 Verify location first":"🤳 Complete face ID first"}
          </p>
        )}
        {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":
          msg.startsWith("🤖")||msg.startsWith("📸")||msg.startsWith("🤳")?"#FFA500":HW.red,
          fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
      </div>

      {aiResult&&(
        <div style={{...s.card,border:`1px solid rgba(207,10,44,.4)`}}>
          <h3 style={{marginBottom:20,color:HW.red}}>🤖 AI Analysis Result</h3>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20}}>
            <div style={{background:HW.surface2,borderRadius:12,padding:16,textAlign:"center"}}>
              <div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:8}}>KPI Score</div>
              <div style={{fontSize:52,fontWeight:800,color:kpiColor(aiResult.kpi_score)}}>{aiResult.kpi_score}</div>
              <div style={{fontSize:11,color:HW.muted}}>out of 100</div>
            </div>
            <div style={{background:HW.surface2,borderRadius:12,padding:16}}>
              <div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:12}}>Task Breakdown</div>
              {aiResult.pie_chart&&<PieChart data={aiResult.pie_chart}/>}
            </div>
          </div>
          {aiResult.summary&&(
            <div style={{background:HW.surface2,borderRadius:10,padding:14,marginBottom:12,borderLeft:`3px solid ${HW.red}`}}>
              <div style={{fontSize:11,color:HW.muted,fontWeight:700,marginBottom:6,textTransform:"uppercase"}}>Day Summary</div>
              <div style={{fontSize:14}}>{aiResult.summary}</div>
            </div>
          )}
          {aiResult.talent_notes&&(
            <div style={{background:HW.surface2,borderRadius:10,padding:14,borderLeft:"3px solid #FFA500"}}>
              <div style={{fontSize:11,color:"#FFA500",fontWeight:700,marginBottom:6,textTransform:"uppercase"}}>🌟 Talent Notes</div>
              <div style={{fontSize:14,lineHeight:1.6}}>{aiResult.talent_notes}</div>
            </div>
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
          <div style={{display:"flex",gap:8}}>
            <button style={{...s.btn,background:"rgba(207,10,44,.15)",color:HW.red}}
              onClick={exportExcel}>📊 Export Excel</button>
            <button style={{...s.btn,background:HW.red,color:HW.white}} onClick={logout}>Sign out</button>
          </div>
        </div>

        {/* Main Tabs */}
        <div style={{display:"flex",gap:4,background:HW.surface2,borderRadius:10,
          padding:4,marginBottom:24}}>
          {["trainees","analytics","okr"].map(tab=>(
            <button key={tab} onClick={()=>{setMgmtTab(tab);setSelected(null);setMsg("");}}
              style={{...s.btn,flex:1,padding:"10px",
                background:mgmtTab===tab?HW.surface:"none",
                color:mgmtTab===tab?HW.text:HW.muted,fontSize:13,borderRadius:7,
                borderBottom:mgmtTab===tab?`2px solid ${HW.red}`:"none"}}>
              {tab==="trainees"?"👥 Trainees":tab==="analytics"?"📊 KPI Analytics":"🎯 OKR Tracking"}
            </button>
          ))}
        </div>

        {/* ── TRAINEES TAB ── */}
        {mgmtTab==="trainees"&&!selected&&(
          <div style={s.card}>
            <h3 style={{marginBottom:16}}>All Trainees
              <span style={{fontSize:13,color:HW.muted,fontWeight:400,marginLeft:8}}>
                — click a row to view profile
              </span>
            </h3>
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
                        onClick={e=>{e.stopPropagation();exportPDF(t);}}>
                        📄 PDF
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,
              fontSize:13,marginTop:12,textAlign:"center"}}>{msg}</p>}
          </div>
        )}

        {/* ── PROFILE VIEW ── */}
        {mgmtTab==="trainees"&&selected&&(
          <div>
            <button style={{...s.btn,background:HW.surface2,color:HW.text,
              marginBottom:20,border:`1px solid ${HW.border}`}}
              onClick={()=>{setSelected(null);setMsg("");}}>
              ← Back to table
            </button>

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
                  <div style={{color:HW.muted,fontSize:13}}>{selected.department} · {selected.assigned_mentor}</div>
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
                          {Math.floor((new Date(selected.quitting_date||new Date())-new Date(selected.joining_date))/(1000*60*60*24))} days
                        </b>
                      </span>
                    )}
                  </div>
                  <span style={{padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,
                    marginTop:6,display:"inline-block",
                    background:statusColors[selected.status]?.bg,color:statusColors[selected.status]?.color}}>
                    {selected.status}
                  </span>
                </div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button style={{...s.btn,background:"rgba(207,10,44,.15)",color:HW.red,fontSize:12}}
                  onClick={()=>exportPDF()}>📄 Export PDF</button>
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

            <div style={{display:"flex",gap:4,background:HW.surface2,borderRadius:10,padding:4,marginBottom:20}}>
              {["timeline","edit","reports","penalties"].map(tab=>(
                <button key={tab} onClick={()=>setProfileTab(tab)}
                  style={{...s.btn,flex:1,padding:"8px",
                    background:profileTab===tab?HW.surface:"none",
                    color:profileTab===tab?HW.text:HW.muted,
                    fontSize:12,borderRadius:7,
                    borderBottom:profileTab===tab?`2px solid ${HW.red}`:"none"}}>
                  {tab==="timeline"?"📅 Timeline":tab==="edit"?"✏️ Edit":
                   tab==="reports"?"📋 Reports":"⚠️ Penalties"}
                </button>
              ))}
            </div>

            {profileTab==="timeline"&&(
              <div style={s.card}>
                <h3 style={{marginBottom:20}}>Activity Timeline</h3>
                {logs.length===0?<p style={{color:HW.muted}}>No activity logged yet.</p>
                  :logs.map((log,i)=>(
                    <div key={log.id} style={{display:"flex",gap:16,marginBottom:20,position:"relative"}}>
                      {i<logs.length-1&&(
                        <div style={{position:"absolute",left:19,top:40,width:2,height:"calc(100% + 4px)",background:HW.border}}/>
                      )}
                      <div style={{width:40,height:40,borderRadius:"50%",background:HW.surface2,
                        border:`2px solid ${HW.border}`,display:"flex",alignItems:"center",
                        justifyContent:"center",fontSize:18,flexShrink:0,zIndex:1}}>
                        {eventIcons[log.event_type]||"📌"}
                      </div>
                      <div style={{flex:1,paddingTop:6}}>
                        <div style={{fontWeight:600,fontSize:14}}>{log.description}</div>
                        {log.old_value&&log.new_value&&(
                          <div style={{fontSize:12,color:HW.muted,marginTop:4}}>
                            <span style={{color:"#f87171"}}>{log.old_value}</span>{" → "}
                            <span style={{color:"#34d399"}}>{log.new_value}</span>
                          </div>
                        )}
                        <div style={{fontSize:11,color:HW.muted,marginTop:4}}>
                          {new Date(log.created_at).toLocaleDateString("en-GB",{
                            day:"numeric",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"
                          })}
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
                      onChange={e=>setSelected({...selected,department:e.target.value,_original:selected._original||{...selected}})}/></div>
                  <div><label style={s.label}>Assigned Mentor</label>
                    <input style={s.input} value={selected.assigned_mentor||""}
                      onChange={e=>setSelected({...selected,assigned_mentor:e.target.value,_original:selected._original||{...selected}})}/></div>
                  <div><label style={s.label}>GPA</label>
                    <input style={s.input} type="number" step="0.01" min="0" max="4" value={selected.gpa||""}
                      onChange={e=>setSelected({...selected,gpa:e.target.value})}/></div>
                  <div><label style={s.label}>Joining Date</label>
                    <input style={s.input} type="date" value={selected.joining_date||""}
                      onChange={e=>setSelected({...selected,joining_date:e.target.value})}/></div>
                </div>
                <button style={{...s.btn,background:HW.red,color:HW.white,marginTop:16}} onClick={saveProfile}>Save Changes</button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:10}}>{msg}</p>}
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
                      <div key={r.id} style={{background:HW.surface2,borderRadius:12,padding:16,marginBottom:14}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                          <div style={{fontWeight:600}}>📅 {r.report_date}</div>
                          <div style={{display:"flex",alignItems:"center",gap:12}}>
                            {r.penalty_applied&&(
                              <span style={{fontSize:11,color:HW.red,background:"rgba(207,10,44,.1)",
                                padding:"2px 8px",borderRadius:10,fontWeight:700}}>⚠️ -{r.penalty_amount}%</span>
                            )}
                            {r.kpi_score&&(
                              <div style={{fontSize:22,fontWeight:800,color:kpiColor(r.kpi_score)}}>
                                {r.kpi_score}<span style={{fontSize:11,color:HW.muted,fontWeight:400}}> KPI</span>
                              </div>
                            )}
                          </div>
                        </div>
                        <div style={{fontSize:11,color:HW.muted,marginBottom:8}}>
                          {r.signin_time||"no sign-in"} ·{" "}
                          <span style={{color:r.attended?"#34d399":HW.red}}>{r.attended?"● Attended":"○ Absent"}</span>
                        </div>
                        {r.excuse_type&&(
                          <div style={{background:"rgba(255,165,0,.08)",borderRadius:8,padding:10,
                            marginBottom:10,borderLeft:"3px solid #FFA500"}}>
                            <div style={{fontSize:11,color:"#FFA500",fontWeight:700,marginBottom:4}}>Excuse: {r.excuse_type}</div>
                            <div style={{fontSize:12}}>{r.excuse_text}</div>
                          </div>
                        )}
                        <div style={{fontSize:13,marginBottom:12,borderLeft:`3px solid ${HW.red}`,paddingLeft:10}}>{r.report_text}</div>
                        {r.face_capture_url&&(
                          <div style={{marginBottom:10}}>
                            <div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>🤳 Face ID</div>
                            <img src={r.face_capture_url} alt="face"
                              style={{width:80,height:80,objectFit:"cover",borderRadius:"50%",border:`2px solid ${HW.red}`}}/>
                          </div>
                        )}
                        {pie&&(
                          <div style={{marginBottom:12}}>
                            <div style={{fontSize:11,color:HW.muted,fontWeight:700,textTransform:"uppercase",marginBottom:8}}>Task Breakdown</div>
                            <PieChart data={pie}/>
                          </div>
                        )}
                        {r.talent_notes&&(
                          <div style={{background:"rgba(207,10,44,.06)",borderRadius:8,padding:10,borderLeft:`3px solid ${HW.red}`}}>
                            <div style={{fontSize:11,color:HW.red,fontWeight:700,marginBottom:4}}>🌟 AI Talent Notes</div>
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
                    <div style={{background:"rgba(207,10,44,.08)",border:`1px solid rgba(207,10,44,.2)`,
                      borderRadius:10,padding:14,marginBottom:16}}>
                      <div style={{fontSize:13,color:HW.red,fontWeight:700}}>
                        Total: {penalties.length} day(s) · {(penalties.length*PENALTY_PCT).toFixed(2)}% deduction
                      </div>
                    </div>
                    {penalties.map(p=>(
                      <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",
                        background:HW.surface2,borderRadius:10,borderLeft:`3px solid ${HW.red}`,padding:12,marginBottom:10}}>
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
          </div>
        )}

        {/* ══ ANALYTICS TAB ══ */}
        {mgmtTab==="analytics"&&(
          <div>
            {/* KPI Summary Cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:16,marginBottom:24}}>
              {[
                {label:"Active Trainees",value:analytics.active.length,icon:"👥",color:HW.red},
                {label:"Avg KPI Score",value:`${analytics.avgKpi}`,icon:"📊",color:"#FFA500"},
                {label:"Attendance Rate",value:`${analytics.attendanceRate}%`,icon:"✅",color:"#34d399"},
                {label:"Total Penalties",value:analytics.totalPenalties,icon:"⚠️",color:"#f87171"},
                {label:"Total Reports",value:allReports.length,icon:"📋",color:HW.red},
                {label:"Departments",value:[...new Set(trainees.map(t=>t.department).filter(Boolean))].length,icon:"🏢",color:"#FFA500"},
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
              {/* KPI Leaderboard */}
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>🏆 KPI Leaderboard</h3>
                {analytics.traineeKpi.length===0
                  ?<p style={{color:HW.muted}}>No data yet.</p>
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
                        <div style={{fontSize:11,color:HW.muted}}>{t.department} · {t.reports} reports</div>
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

              {/* Department Performance */}
              <div style={s.card}>
                <h3 style={{marginBottom:16}}>🏢 Department KPI Performance</h3>
                {analytics.deptData.length===0
                  ?<p style={{color:HW.muted}}>No data yet.</p>
                  :<BarChart data={analytics.deptData}/>
                }
                <div style={{marginTop:16}}>
                  {analytics.deptData.map((d,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",
                      alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${HW.border}`}}>
                      <span style={{fontSize:13,color:HW.muted}}>{d.label}</span>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:80,height:6,background:HW.border,borderRadius:10,overflow:"hidden"}}>
                          <div style={{height:"100%",background:HW.red,borderRadius:10,
                            width:`${d.value}%`}}/>
                        </div>
                        <span style={{fontSize:13,fontWeight:700,color:kpiColor(d.value),minWidth:30}}>
                          {d.value}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* At Risk Trainees */}
            {analytics.atRisk.length>0&&(
              <div style={{...s.card,border:`1px solid rgba(248,113,113,.3)`}}>
                <h3 style={{marginBottom:16,color:"#f87171"}}>
                  ⚠️ At Risk Trainees
                  <span style={{fontSize:13,color:HW.muted,fontWeight:400,marginLeft:8}}>
                    — KPI below 60 or more than 2 penalties
                  </span>
                </h3>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:12}}>
                  {analytics.atRisk.map(t=>(
                    <div key={t.id} style={{background:HW.surface2,borderRadius:12,padding:16,
                      borderLeft:`3px solid #f87171`}}>
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

        {/* ══ OKR TAB ══ */}
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

            {/* Add OKR Form */}
            {showAddOkr&&(
              <div style={{...s.card,border:`1px solid rgba(207,10,44,.3)`}}>
                <h4 style={{marginBottom:16,color:HW.red}}>Add New OKR</h4>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
                  <div><label style={s.label}>Department</label>
                    <input style={s.input} placeholder="e.g. Engineering" value={newOkr.department}
                      onChange={e=>setNewOkr({...newOkr,department:e.target.value})}/></div>
                  <div><label style={s.label}>Due Date</label>
                    <input style={s.input} type="date" value={newOkr.due_date}
                      onChange={e=>setNewOkr({...newOkr,due_date:e.target.value})}/></div>
                  <div style={{gridColumn:"1/-1"}}><label style={s.label}>Objective</label>
                    <input style={s.input} placeholder="e.g. Improve Technical Skills" value={newOkr.objective}
                      onChange={e=>setNewOkr({...newOkr,objective:e.target.value})}/></div>
                  <div style={{gridColumn:"1/-1"}}><label style={s.label}>Key Result</label>
                    <input style={s.input} placeholder="e.g. Complete 20 lab sessions" value={newOkr.key_result}
                      onChange={e=>setNewOkr({...newOkr,key_result:e.target.value})}/></div>
                  <div><label style={s.label}>Target</label>
                    <input style={s.input} type="number" value={newOkr.target}
                      onChange={e=>setNewOkr({...newOkr,target:parseFloat(e.target.value)})}/></div>
                  <div><label style={s.label}>Unit</label>
                    <select style={s.input} value={newOkr.unit}
                      onChange={e=>setNewOkr({...newOkr,unit:e.target.value})}>
                      <option value="%">% (Percentage)</option>
                      <option value="sessions">Sessions</option>
                      <option value="reports">Reports</option>
                      <option value="tasks">Tasks</option>
                      <option value="reviews">Reviews</option>
                      <option value="workflows">Workflows</option>
                    </select>
                  </div>
                </div>
                <button style={{...s.btn,background:HW.red,color:HW.white,marginTop:16}} onClick={addOkr}>
                  Add OKR
                </button>
                {msg&&<p style={{color:msg.startsWith("✅")?"#34d399":HW.red,fontSize:13,marginTop:10}}>{msg}</p>}
              </div>
            )}

            {/* OKR Summary */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:16,marginBottom:20}}>
              {[
                {label:"Total OKRs",value:okrs.length,color:HW.red},
                {label:"On Track (≥80%)",value:okrs.filter(o=>(o.current/o.target)>=.8).length,color:"#34d399"},
                {label:"In Progress",value:okrs.filter(o=>(o.current/o.target)>=.5&&(o.current/o.target)<.8).length,color:"#FFA500"},
                {label:"Behind (<50%)",value:okrs.filter(o=>(o.current/o.target)<.5).length,color:"#f87171"},
              ].map((s2,i)=>(
                <div key={i} style={{background:HW.surface,border:`1px solid ${HW.border}`,
                  borderRadius:14,padding:16,textAlign:"center",borderTop:`3px solid ${s2.color}`}}>
                  <div style={{fontSize:28,fontWeight:800,color:s2.color}}>{s2.value}</div>
                  <div style={{fontSize:12,color:HW.muted,marginTop:4}}>{s2.label}</div>
                </div>
              ))}
            </div>

            {/* OKR List by Department */}
            {[...new Set(okrs.map(o=>o.department))].map(dept=>(
              <div key={dept} style={s.card}>
                <h4 style={{marginBottom:16,color:HW.red}}>🏢 {dept}</h4>
                {okrs.filter(o=>o.department===dept).map(okr=>(
                  <OKRBar key={okr.id} okr={okr} isManager={true} onUpdate={updateOkr}/>
                ))}
              </div>
            ))}

            {okrs.length===0&&(
              <div style={{...s.card,textAlign:"center",padding:40}}>
                <div style={{fontSize:40,marginBottom:12}}>🎯</div>
                <p style={{color:HW.muted}}>No OKRs yet. Click "+ Add OKR" to create your first one!</p>
              </div>
            )}
          </div>
        )}
      </div>
    );
  }
}