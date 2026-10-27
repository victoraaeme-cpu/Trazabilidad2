const $=id=>document.getElementById(id);
const deb=(f,ms=200)=>{let t;return(...a)=>{clearTimeout(t);t=setTimeout(()=>f(...a),ms)}};
// --- Carga diferida de librerías pesadas ---
const _libs={};
function cargarScript(src){return _libs[src]||(_libs[src]=new Promise((ok,ko)=>{const e=document.createElement("script");e.src=src;e.onload=ok;e.onerror=()=>{delete _libs[src];ko(new Error("No se pudo cargar "+src))};document.head.appendChild(e)}))}
const cargarXLSX=()=>window.XLSX?Promise.resolve():cargarScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js").catch(()=>{throw new Error("No se pudo cargar el lector de Excel. Revise la conexión e inténtelo de nuevo.")});
const cargarJsQR=()=>window.jsQR?Promise.resolve():cargarScript("js/vendor/jsqr.js").catch(e=>{toast("No se pudo cargar el lector QR alterno. Revise la conexión.",1);throw e});
// --- Índices por código y caché de existencia ---
const _ex=new WeakMap();
function existencia(m){let e=_ex.get(m);if(!e){e=existenciaCalc(m);_ex.set(m,e)}return e}
const _ix={};
function indice(clave,lista){const c=_ix[clave];if(!c||c.l!==lista||c.n!==lista.length)_ix[clave]={l:lista,n:lista.length,m:new Map(lista.map(x=>[x.codigo,x]))};return _ix[clave].m}
const ubiPorCodigo=c=>indice("u",ubicaciones).get(c);
const almPorCodigo=c=>indice("a",almacenes).get(c);
// --- Diálogo de confirmación (con texto obligatorio para acciones destructivas) ---
function confirmar(o){if(typeof o==="string")o={mensaje:o};
 return new Promise(res=>{const d=$("dlgConfirmar"),inp=$("dlgTexto"),ok=$("dlgOk"),no=$("dlgNo");
  $("dlgTitulo").textContent=o.titulo||"Confirmar";$("dlgMsg").textContent=o.mensaje||"";
  $("dlgEscribirWrap").hidden=!o.escribir;$("dlgEscribirLbl").textContent=o.escribir?`Escribe «${o.escribir}» para continuar`:"";inp.value="";
  ok.textContent=o.boton||"Confirmar";ok.className=o.peligro?"bd":"bp";ok.disabled=!!o.escribir;
  const fin=v=>{d.close();inp.oninput=ok.onclick=no.onclick=d.oncancel=null;res(v)};
  inp.oninput=()=>{ok.disabled=inp.value.trim().toUpperCase()!==String(o.escribir).toUpperCase()};
  ok.onclick=()=>fin(true);no.onclick=()=>fin(false);d.oncancel=e=>{e.preventDefault();fin(false)};
  d.showModal();(o.escribir?inp:no).focus()})}
// --- Indicador de conexión + reconexión automática ---
let rtOk=false;
function pintarConexion(){const el=$("conEstado");if(!el)return;const on=navigator.onLine;el.className="chip con "+(!on?"off":rtOk?"ok":"wait");el.textContent=!on?"● Sin conexión":rtOk?"● En línea":"● Conectando…";el.title=!on?"Sin internet: los cambios no se guardarán":rtOk?"Sincronización en tiempo real activa":"Conectando con el servidor…"}
function marcarRT(est){rtOk=est==="SUBSCRIBED";pintarConexion();if(est==="CHANNEL_ERROR"||est==="TIMED_OUT"||est==="CLOSED")setTimeout(()=>{if(esStaff&&!canalRT)suscribir()},5000)}
let _rec=false,_pend=false;
async function recargarTodo(){if(_rec){_pend=true;return}_rec=true;try{await Promise.allSettled([cargar(),cargarUbicaciones()])}finally{_rec=false;if(_pend){_pend=false;recargarTodo()}}}
addEventListener("online",()=>{pintarConexion();if(esStaff)recargarTodo()});addEventListener("offline",pintarConexion);
// --- Tema claro/oscuro ---
(function(){let t=null;try{t=localStorage.getItem("tema_ui")}catch{}if(t)document.documentElement.dataset.theme=t;
 addEventListener("DOMContentLoaded",()=>{pintarConexion();$("bTema").onclick=()=>{const osc=matchMedia("(prefers-color-scheme:dark)").matches,act=document.documentElement.dataset.theme||(osc?"dark":"light"),nv=act==="dark"?"light":"dark";document.documentElement.dataset.theme=nv;try{localStorage.setItem("tema_ui",nv)}catch{}}})})();

const K="trazabilidad_materiales_v2";
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const hoy=()=>new Date(Date.now()-new Date().getTimezoneOffset()*6e4).toISOString().slice(0,10);
const fmt=n=>Number(n).toLocaleString("es-MX",{maximumFractionDigits:3});
function totalTxt(m){return m.esMP&&m.um==="PZ"&&m.cantidad>0&&m.longitud>0?`${fmt(m.cantidad*m.longitud)} ${m.longU||m.longitudUnidad||"mm"}`:""}
let mats=[],actual=null,orden={k:null,asc:true},esStaff=false,esAdmin=false,almacenes=[],ubicaciones=[],ubicacionesAlias=[],movimientosUbicacion=[],ajustesInventarioError="",almacenesError="",ubicacionesError="",ubicacionesAliasError="",movimientosError="",ubicacionDestino=null,asignacionItems=[];
const seleccionados=new Set();

// --- Supabase ---
const SB_URL="https://olbcepxgzqtmiqdqpbha.supabase.co",SB_KEY="sb_publishable_kdySoe6yEpOSmx4-zEOk-w_EHIMnaEC";
const sb=window.supabase?window.supabase.createClient(SB_URL,SB_KEY):null;
function normalizarMaterial(m){
  const pesoU=m.pesoU||m.pesoUnidad||(["LB","KG"].includes(m.um)?m.um:m.peso!=null&&!m.pesoUnidadPendiente?"LB":null);
  const tipoRegistro=/^(?:MP|MATERIA PRIMA)$/.test(String(m.tipoRegistro??"").toUpperCase())||m.esMP?"MP":/^CONSUMIBLES?$/.test(String(m.tipoRegistro??"").trim().toUpperCase())?"CONSUMIBLES":"GENERAL";
  const seccionMaterial=normalizarSeccion(m.seccionMaterial)||(m.seccion==="cuadrado"?"SQ":m.seccion==="redondo"?"RND":"");
  return{...m,um:"PZ",pesoU,tipoRegistro,esMP:tipoRegistro==="MP",seccionMaterial};
}
const nombreTipoRegistro=tipo=>({MP:"Materia prima",GENERAL:"Material general",CONSUMIBLES:"Consumibles"}[tipo]||"Material general");
const matData=m=>{const{despachos,ajustesInventario,...r}=m;return r};
const desData=d=>{const{t,piezas,...r}=d;return r};
const desFila=(codigo,d)=>({codigo,t:d.t,data:desData(d),piezas:d.piezas||[]});
const desDeFila=r=>({...r.data,t:Number(r.t),...(r.piezas&&r.piezas.length?{piezas:r.piezas}:{})});
function errDB(e,msg){console.error(e);const t=String((e&&e.message)||e||"");
  toast(/EXISTENCIA_CAMBIO_DURANTE_CONTEO/.test(t)?"La existencia cambió mientras realizabas el conteo. Actualiza la consulta e inténtalo de nuevo.":/PIEZA_YA_SURTIDA/.test(t)?"⚠ Una de esas piezas ya fue surtida desde otro equipo":/duplicate|23505/.test(t+((e&&e.code)||""))?"⚠ El código ya existe, use otro.":`${msg} (${t.slice(0,90)})`,1)}
async function leerTodo(tabla,cols,ord){const out=[];
  for(let i=0;;i+=1000){let q=sb.from(tabla).select(cols);ord.forEach(o=>{q=q.order(o)});const{data,error}=await q.range(i,i+999);if(error)throw error;out.push(...data);if(data.length<1000)break}
  return out}
// Recarga los datos y conserva la edición, el QR y los renglones pendientes.
async function cargar(){
  const[materialesResultado,despachosResultado,ajustesResultado]=await Promise.allSettled([
    leerTodo("materiales","codigo,data",["codigo"]),
    leerTodo("despachos","codigo,t,data,piezas",["t","id"]),
    leerTodo("ajustes_inventario_ciclico","id,codigo_material,ubicacion_codigo,ubicacion_nombre,existencia_anterior,existencia_contada,diferencia,unidad,motivo,usuario,created_at",["id"])
  ]);
  if(materialesResultado.status==="rejected")throw materialesResultado.reason;
  if(despachosResultado.status==="rejected")throw despachosResultado.reason;
  const A=materialesResultado.value,B=despachosResultado.value;
  ajustesInventarioError=ajustesResultado.status==="rejected"
    ?(/PGRST205|42P01|does not exist|schema cache/i.test(String(ajustesResultado.reason?.message||ajustesResultado.reason))?"Falta aplicar o actualizar la migración supabase_almacenes.sql para habilitar los conteos cíclicos.":`No se pudieron cargar los conteos cíclicos: ${String(ajustesResultado.reason?.message||ajustesResultado.reason)}`)
    :"";
  if(ajustesResultado.status==="rejected")console.error(ajustesResultado.reason);
  const ajustesPor={};
  if(ajustesResultado.status==="fulfilled")ajustesResultado.value.forEach(a=>{(ajustesPor[a.codigo_material]=ajustesPor[a.codigo_material]||[]).push(a)});
  const por={};B.forEach(r=>{(por[r.codigo]=por[r.codigo]||[]).push(desDeFila(r))});
  const edI=$("editIndex").value,edCod=edI!==""&&mats[+edI]?mats[+edI].codigo:null;
  const act=actual,actAntes=act?JSON.stringify(act):"";
  mats=A.map(r=>({...normalizarMaterial(r.data),codigo:r.codigo,despachos:por[r.codigo]||[],ajustesInventario:ajustesPor[r.codigo]||[]}));
  if(edCod){const i=mats.findIndex(x=>x.codigo===edCod);if(i<0)resetForm();else $("editIndex").value=i}
  surItems=surItems.filter(item=>mats.some(x=>x.codigo.toLowerCase()===item.codigo.toLowerCase()));
  surItems.forEach(item=>{const m=mats.find(x=>x.codigo.toLowerCase()===item.codigo.toLowerCase());if(item.piezas?.some(n=>piezasSurt(m).has(n)))item.errorPieza=true});
  render();
  if(!$("vInv").hidden)renderInventario();
  if(!$("vSur").hidden){renderHist();renderSur()}
  if(act){const a=mats.find(x=>x.codigo===act.codigo);if(a){if(JSON.stringify(a)!==actAntes)mostrar(a);else actual=a}}
}
async function cargarUbicaciones(){
  const resultados=await Promise.allSettled([
    leerTodo("almacenes","codigo,nombre,clasificacion,created_at",["nombre"]),
    leerTodo("ubicaciones","codigo,nombre,almacen_codigo,created_at",["nombre"]),
    leerTodo("movimientos_ubicacion","id,folio,codigo_material,ubicacion_anterior_codigo,ubicacion_anterior,ubicacion_nueva_codigo,ubicacion_nueva,ubicacion_anterior_almacen_codigo,ubicacion_nueva_almacen_codigo,existencia,created_at",["id"]),
    leerTodo("ubicaciones_codigos_alias","codigo_alias,ubicacion_codigo,created_at",["codigo_alias"])
  ]);
  const mensajeError=(resultado,nombre)=>{
    if(resultado.status==="fulfilled")return"";
    const detalle=String(resultado.reason?.message||resultado.reason||"Error desconocido");
    if(/PGRST205|42P01|does not exist|schema cache/i.test(detalle))return`Falta aplicar o actualizar la migración: no se pudo consultar ${nombre}.`;
    return`No se pudo consultar ${nombre}: ${detalle}`;
  };
  almacenesError=mensajeError(resultados[0],"almacenes");
  ubicacionesError=mensajeError(resultados[1],"ubicaciones");
  movimientosError=mensajeError(resultados[2],"el historial de movimientos");
  ubicacionesAliasError=mensajeError(resultados[3],"los códigos anteriores de QR de ubicaciones");
  almacenes=resultados[0].status==="fulfilled"?resultados[0].value:[];
  ubicaciones=resultados[1].status==="fulfilled"?resultados[1].value:[];
  movimientosUbicacion=resultados[2].status==="fulfilled"?resultados[2].value:[];
  ubicacionesAlias=resultados[3].status==="fulfilled"?resultados[3].value:[];
  resultados.forEach(resultado=>{if(resultado.status==="rejected")console.error(resultado.reason)});
  if(!$("ubiCodigo").value)siguienteCodigoUbicacion();
  if(!$("almacenCodigo").value)siguienteCodigoAlmacen();
  renderAlmacenes();renderUbicaciones();renderHistUbicaciones();renderInventario();actualizarDatalistUbicaciones();
}
async function subirLista(ms,ds){
  if(!esAdmin)throw new Error("Solo el administrador puede restaurar información.");
  for(let i=0;i<ms.length;i+=200){const{error}=await sb.from("materiales").upsert(ms.slice(i,i+200).map(m=>({codigo:m.codigo,data:matData(m)})),{onConflict:"codigo"});if(error)throw error}
  const filas=ds.map(x=>desFila(x.codigo,x.d));
  for(let i=0;i<filas.length;i+=200){const{error}=await sb.from("despachos").upsert(filas.slice(i,i+200),{onConflict:"codigo,t",ignoreDuplicates:true});if(error)throw error}
}
let rtT=null,canalRT=null;
function dejarCanal(){const c=canalRT;canalRT=null;if(!sb||!c)return;try{sb.removeChannel(c).catch(()=>{})}catch{}}
function suscribir(){
  if(!sb||canalRT)return;
  const rt=()=>{clearTimeout(rtT);rtT=setTimeout(recargarTodo,600)};
  canalRT=sb.channel("cambios").on("postgres_changes",{event:"*",schema:"public",table:"materiales"},rt).on("postgres_changes",{event:"*",schema:"public",table:"despachos"},rt).on("postgres_changes",{event:"*",schema:"public",table:"almacenes"},rt).on("postgres_changes",{event:"*",schema:"public",table:"ubicaciones"},rt).on("postgres_changes",{event:"*",schema:"public",table:"ubicaciones_codigos_alias"},rt).on("postgres_changes",{event:"*",schema:"public",table:"movimientos_ubicacion"},rt).on("postgres_changes",{event:"*",schema:"public",table:"ajustes_inventario_ciclico"},rt).subscribe(estado=>{marcarRT(estado);if((estado==="CHANNEL_ERROR"||estado==="TIMED_OUT")&&canalRT){const c=canalRT;canalRT=null;try{sb.removeChannel(c).catch(()=>{})}catch{}}});
}
// --- Sesión ---
function lgMostrar(msg){$("lgCarga").hidden=true;$("lgF").hidden=false;$("lgMsg").hidden=!msg;$("lgMsg").textContent=msg||""}
function limpiarSesionLocal(){try{Object.keys(localStorage).filter(k=>k.indexOf("sb-")===0&&k.indexOf("-auth-token")>0).forEach(k=>localStorage.removeItem(k))}catch{}}
function errorSesion(e){const c=Number(e?.status||e?.code||0);return c===401||c===403||/jwt|token|session|sesi[oó]n/i.test(String(e?.message||""))}
async function cerrarSesion(){try{const{error}=await sb.auth.signOut();if(error)throw error}catch(e){console.warn("No se pudo revocar la sesión en el servidor; se cierra localmente.",e)}limpiarSesionLocal();salirUI()}
function actualizarVisibilidadAdmin(){["bJson","bCsv","bImp","bExcel","bCsvS","bUbiExportar","ubicacionesImportar","btnGuardar"].forEach(id=>$(id).hidden=!esAdmin);$("registroAdminAviso").hidden=esAdmin}
function salirUI(){esStaff=false;esAdmin=false;actualizarVisibilidadAdmin();seleccionados.clear();$("rolUsuario").hidden=true;$("conteoEstado").hidden=true;document.body.classList.remove("auth");dejarCanal();camOff();if($("conteoDialog").open)$("conteoDialog").close();mats=[];surItems=[];almacenes=[];ubicaciones=[];ubicacionesAlias=[];movimientosUbicacion=[];ajustesInventarioError="";almacenesError="";ubicacionesError="";ubicacionesAliasError="";movimientosError="";ubicacionDestino=null;asignacionItems=[];actual=null;render();lgMostrar()}
let entrando=false;
async function entrar(){
  if(!sb||entrando)return false;
  entrando=true;
  try{
    const admin=await sb.rpc("is_admin");
    if(admin.error){const expirada=errorSesion(admin.error);await cerrarSesion();lgMostrar(expirada?"Tu sesión expiró. Vuelve a iniciar sesión.":"No se pudieron verificar los permisos del administrador. Ejecute la migración actualizada supabase_almacenes.sql.");return false}
    const r=await sb.rpc("is_staff");
    if((r.error||r.data!==true)&&admin.data!==true){const expirada=errorSesion(r.error);await cerrarSesion();lgMostrar(expirada?"Tu sesión expiró. Vuelve a iniciar sesión.":"Este usuario no tiene autorización. Pida al administrador que agregue su correo.");return false}
    esStaff=true;esAdmin=admin.data===true;document.body.classList.add("auth");
    $("rolUsuario").textContent=esAdmin?"Administrador":"Personal";$("rolUsuario").hidden=false;
    $("bDelSel").hidden=$("bDelAll").hidden=$("selectAll").hidden=!esAdmin;
    actualizarVisibilidadAdmin();
    try{await cargar()}catch(e){errDB(e,"No se pudo cargar los datos")}
    await cargarUbicaciones();
    suscribir();await migrarLocal();return true;
  }finally{entrando=false}
}
// Datos que el navegador tenía guardados antes de usar Supabase
async function migrarLocal(){
  if(mats.length||sessionStorage.getItem("mig_omitido"))return;
  let loc=[];try{if(localStorage.getItem(K+"_migrado"))return;loc=JSON.parse(localStorage.getItem(K)||localStorage.getItem("materiales_v1")||"[]")}catch{}
  loc=(Array.isArray(loc)?loc:[]).filter(m=>m&&m.codigo);if(!loc.length)return;
  if(!esAdmin){toast("La restauración de datos guardados en este navegador requiere al administrador.",1);return}
  if(!await confirmar(`Se encontraron ${loc.length} registro(s) guardados en este navegador (de antes de usar Supabase).\n\n¿Subirlos ahora a la base de datos compartida?`)){try{sessionStorage.setItem("mig_omitido","1")}catch{}return}
  loc.forEach(m=>{Object.assign(m,normalizarMaterial(m));m.codigo=String(m.codigo);m.esMP=!!m.esMP;m.cantidad=Number(m.cantidad)||0;m.despachos=Array.isArray(m.despachos)?m.despachos.filter(d=>d&&typeof d==="object"&&d.t!=null):[]});
  try{await subirLista(loc,loc.flatMap(m=>m.despachos.map(d=>({codigo:m.codigo,d}))));try{localStorage.setItem(K+"_migrado","1")}catch{}await cargar();toast(`${loc.length} registro(s) subidos a Supabase ✔`)}
  catch(e){errDB(e,"No se pudo subir")}
}
$("lgF").addEventListener("submit",async e=>{
  e.preventDefault();if($("lgBtn").disabled)return;
  $("lgMsg").hidden=true;$("lgBtn").disabled=true;
  try{
    const{error}=await sb.auth.signInWithPassword({email:$("lgMail").value.trim(),password:$("lgPass").value});
    if(error){lgMostrar("Correo o contraseña incorrectos.");return}
    $("lgPass").value="";await entrar();
  }finally{$("lgBtn").disabled=false}
});
$("bSalir").onclick=()=>{cerrarSesion()};
async function arrancar(){
  const ficha=/#(?:[^#]*&)?[dc]=/.test(location.hash);if(ficha)document.body.classList.add("ficha");
  if(!sb){lgMostrar("No se pudo cargar Supabase. Revise su conexión a internet y recargue la página.");if(ficha)iniciarFicha();return}
  sb.auth.onAuthStateChange(ev=>{if(ev==="SIGNED_OUT")salirUI()});
  let ses=null;try{ses=(await sb.auth.getSession()).data.session}catch{}
  if(ses)await entrar();else lgMostrar();
  if(ficha)iniciarFicha();
}

function toast(msg,err){const t=$("toast");t.textContent=msg;t.classList.toggle("err",!!err);t.classList.add("show");clearTimeout(toast.h);toast.h=setTimeout(()=>t.classList.remove("show"),err?5000:2400)}

// --- Formulario ---
const CAMPOS_EXTRA=[["tipoMaterial","tipoMaterial"],["especificacionMedida","especificacionMedida"],["longitudIndividual","longitudIndividual"],["manufactura","manufactura"],["numeroHeat","numeroHeat"],["numeroLoteMP","numeroLoteMP"],["numeroCertificado","numeroCertificado"],["totalMM","totalMM"],["peso","peso"],["ordenMaterial","ordenMaterial"],["poNumber","poNumber"]];
const numeroCampo=id=>$(id).value===""?null:Number($(id).value);
function recalcularTotalMM(origen){
  const naTotal=document.querySelector('input[data-na="totalMM"]').checked;
  if(naTotal)return;
  const longInput=$("longitudIndividual"),qtyInput=$("cantidad");
  if(longInput.value===""){
    if(origen==="longitudIndividual")$("totalMM").value="";
    return;
  }
  const longitud=Number(longInput.value),cantidad=Number(qtyInput.value);
  if(Number.isFinite(longitud)&&Number.isFinite(cantidad)&&longitud>=0&&cantidad>=0)
    $("totalMM").value=String(longitud*cantidad);
}
document.querySelectorAll("[data-na]").forEach(c=>c.addEventListener("change",()=>{
  const input=$(c.dataset.na),wrap=input.closest(".na-field").querySelector(".na-value-wrap")||input;
  wrap.hidden=c.checked;if(c.checked)input.value="";
  if(c.dataset.na==="seccionMaterial")actualizarCamposSeccion();
  if(c.dataset.na==="longitudIndividual")recalcularTotalMM("longitudIndividual");
  if(c.dataset.na==="totalMM"&&!c.checked)recalcularTotalMM();
}));
$("longitudIndividual").addEventListener("input",()=>recalcularTotalMM("longitudIndividual"));
$("cantidad").addEventListener("input",()=>recalcularTotalMM("cantidad"));
function actualizarCamposSeccion(){
  const seccion=$("seccionMaterial").value;
  $("dimensionesLineales").hidden=seccion!=="SQ"&&seccion!=="FLAT";
  $("dimensionRedonda").hidden=seccion!=="RND";
}
$("seccionMaterial").addEventListener("change",actualizarCamposSeccion);
$("btnAuto").onclick=()=>{let n=0;mats.forEach(m=>{const x=/^MAT-(\d+)$/i.exec(m.codigo);if(x)n=Math.max(n,+x[1])});$("codigo").value="MAT-"+String(n+1).padStart(4,"0")};

function establecerFormularioAbierto(abierto){$("f").hidden=!abierto;$("tituloForm").setAttribute("aria-expanded",String(abierto))}
$("tituloForm").addEventListener("click",()=>establecerFormularioAbierto($("f").hidden));
$("tituloForm").addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();establecerFormularioAbierto($("f").hidden)}});
function resetForm(){$("f").reset();document.querySelectorAll("[data-na]").forEach(c=>{c.checked=false;$(c.dataset.na).closest(".na-field").querySelector(".na-value-wrap")?.removeAttribute("hidden");$(c.dataset.na).hidden=false});CAMPOS_EXTRA.forEach(([,id])=>$(id).hidden=false);actualizarCamposSeccion();$("editIndex").value="";$("fechaRecibo").value=hoy();$("tituloForm").textContent="Registro de material";$("btnGuardar").textContent="💾 Guardar y generar QR";$("btnCancelar").hidden=true;$("detTraza").open=false;establecerFormularioAbierto(false)}

$("f").addEventListener("invalid",()=>{$("detTraza").open=true},true);
$("f").addEventListener("submit",async e=>{
  e.preventDefault();
  if(!esAdmin){toast("Solo el administrador puede registrar o editar materiales.",1);return}
  recalcularTotalMM();
  const idx=$("editIndex").value;
  const noAplica=Object.fromEntries(CAMPOS_EXTRA.map(([key,id])=>[key,document.querySelector(`input[data-na="${id}"]`).checked]));
  noAplica.seccionMaterial=document.querySelector('input[data-na="seccionMaterial"]').checked;
  const extra=Object.fromEntries(CAMPOS_EXTRA.map(([key,id])=>[key,noAplica[key]?(["totalMM","longitudIndividual","peso"].includes(key)?null:""): (["totalMM","longitudIndividual","peso"].includes(key)?($(id).value===""?null:Number($(id).value)):$(id).value.trim())]));
  const tipoMaterial=noAplica.tipoMaterial?"":$("tipoMaterial").value.trim(),tipoRegistro=$("tipoRegistro").value;
  const anterior=idx!==""?mats[idx]:null;
  const m={codigo:$("codigo").value.trim(),descripcion:$("descripcion").value.trim(),lote:extra.numeroLoteMP||"",proveedor:"",
    um:$("um").value,cantidad:parseFloat($("cantidad").value),ubicacion:$("ubicacion").value.trim(),fechaRecibo:$("fechaRecibo").value,
    fechaModificacion:new Date().toLocaleString("es-MX"),tipoRegistro,esMP:tipoRegistro==="MP",seccionMaterial:noAplica.seccionMaterial?"":$("seccionMaterial").value,
    ubicacionCodigo:anterior&&$("ubicacion").value.trim()===anterior.ubicacion?anterior.ubicacionCodigo||null:null,
    ubicacionAlmacenCodigo:anterior&&$("ubicacion").value.trim()===anterior.ubicacion?anterior.ubicacionAlmacenCodigo||null:null,
    ubicacionAlmacenNombre:anterior&&$("ubicacion").value.trim()===anterior.ubicacion?anterior.ubicacionAlmacenNombre||null:null,
    peso:extra.peso,pesoU:extra.peso==null?null:$("pesoU").value,pesoUnidadPendiente:extra.peso!=null&&!$("pesoU").value,
    longitud:anterior?.longitud??null,longU:anterior?.longU||anterior?.longitudUnidad||null,
    seccion:tipoRegistro==="MP"?({SQ:"cuadrado",FLAT:"cuadrado",RND:"redondo"}[$("seccionMaterial").value]||null):null,
    secU:tipoRegistro==="MP"?"mm":null,
    diametro:tipoRegistro==="MP"&&$("seccionMaterial").value==="RND"?numeroCampo("diametroMaterial"):null,
    alto:tipoRegistro==="MP"&&["SQ","FLAT"].includes($("seccionMaterial").value)?numeroCampo("altoMaterial"):null,
    ancho:tipoRegistro==="MP"&&["SQ","FLAT"].includes($("seccionMaterial").value)?numeroCampo("anchoMaterial"):null,
    ...extra,noAplica};
  m.despachos=anterior?(anterior.despachos||[]):[];
  const dup=mats.findIndex(x=>x.codigo.toLowerCase()===m.codigo.toLowerCase());
  if(dup!==-1&&String(dup)!==idx){toast("⚠ El código ya existe, use otro.",1);return}
  const oldCod=idx!==""?mats[idx].codigo:null;$("btnGuardar").disabled=true;
  let registrandoUbicacion=true,ubicacionCreada=false,ubicacionCorregida=false;
  try{
    const destino=await resolverUbicacionRegistro(m.ubicacion,tipoRegistro,anterior?.ubicacionCodigo);
    registrandoUbicacion=false;
    if(destino){m.ubicacion=destino.ubicacion.nombre;m.ubicacionCodigo=destino.ubicacion.codigo;m.ubicacionAlmacenCodigo=destino.ubicacion.almacen_codigo;m.ubicacionAlmacenNombre=destino.almacen.nombre;ubicacionCreada=destino.creada;ubicacionCorregida=!!destino.corregida}
    if(idx!==""){
      const{error}=await sb.from("materiales").update({codigo:m.codigo,data:matData(m)}).eq("codigo",oldCod);if(error)throw error;
    }else{const{error}=await sb.from("materiales").insert({codigo:m.codigo,data:matData(m)});if(error)throw error}
  }catch(err){errDB(err,registrandoUbicacion?"No se pudo registrar la ubicación":ubicacionCreada?"No se pudo guardar el material; la ubicación quedó registrada":ubicacionCorregida?"No se pudo guardar el material; el código de ubicación y sus materiales anteriores ya se corrigieron":"No se pudo guardar");$("btnGuardar").disabled=false;await cargar().catch(()=>{});return}
  $("btnGuardar").disabled=false;await cargar().catch(()=>{});
  toast(`${idx!==""?"Material actualizado":"Material guardado"}${ubicacionCreada?" · ubicación agregada a Ubicaciones y QR":ubicacionCorregida?" · código de ubicación corregido a "+m.ubicacionCodigo:""} ✔`);
  const m2=mats.find(x=>x.codigo===m.codigo)||m;resetForm();render();try{mostrar(m2)}catch(err){console.error(err);toast("Guardado, pero falló el QR",1)}
});
$("btnCancelar").onclick=()=>{resetForm();toast("Edición cancelada")};

// --- QR ---
function dimsTxt(m){
  if(!m.esMP)return"";
  const seccion=m.seccionMaterial|| (m.seccion==="cuadrado"?"SQ":"RND");
  if(!seccion)return"";
  const etiqueta=seccion==="FLAT"?"FLAT":seccion==="SQ"?"CUADRADO":"REDONDO";
  const medidas=seccion==="RND"?m.diametro!=null?` ⌀ ${m.diametro} ${m.secU||m.diametroUnidad||"mm"}`:"":m.alto!=null||m.ancho!=null?` ${m.alto??"—"} × ${m.ancho??"—"} ${m.secU||"mm"} (alto×ancho)`:"";
  const largo=m.longitud>0?`, L: ${m.longitud} ${m.longU||m.longitudUnidad||"mm"}`:"";
  return `${etiqueta}${medidas}${largo}`;
}
function seccionEspecificacion(m){
  const seccion=m.seccionMaterial|| (m.seccion==="cuadrado"?"SQ":m.seccion==="redondo"?"RND":"");
  const etiqueta=seccion==="SQ"?"SQ":seccion==="FLAT"?"FLAT":seccion==="RND"?"RND":"";
  const especificacion=String(m.especificacionMedida||"").trim();
  const dimension=etiqueta&&especificacion?`${etiqueta} · ${especificacion}`:especificacion||dimsTxt(m);
  return dimension;
}
const BASE="https://shardd89.github.io/TrazabilidadQR/";
const b64e=t=>btoa(encodeURIComponent(t).replace(/%([0-9A-F]{2})/g,(_,h)=>String.fromCharCode(parseInt(h,16)))).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");
const b64d=t=>{t=t.replace(/-/g,"+").replace(/_/g,"/");while(t.length%4)t+="=";return decodeURIComponent(Array.from(atob(t),c=>"%"+c.charCodeAt(0).toString(16).padStart(2,"0")).join(""))};
function payload(m){
  const o={c:m.codigo,d:m.descripcion,l:m.lote,p:m.proveedor,u:"PZ",q:m.cantidad,m:m.esMP?1:0,ub:m.ubicacion,fr:m.fechaRecibo,fm:m.fechaModificacion};
  if(m.esMP)Object.assign(o,{lo:m.longitud,lu:m.longU||m.longitudUnidad,s:m.seccion,su:m.secU||m.diametroUnidad,di:m.diametro,al:m.alto,an:m.ancho});
  if(m.peso!=null){o.pe=m.peso;o.pu=m.pesoUnidadPendiente?"?":m.pesoU||m.pesoUnidad}
  const extras={tipoRegistro:"tr",seccionMaterial:"sm",tipoMaterial:"tm",especificacionMedida:"em",longitudIndividual:"li",manufactura:"mf",numeroHeat:"nh",numeroLoteMP:"nl",numeroCertificado:"nc",totalMM:"tmm",ordenMaterial:"ord",poNumber:"po"};
  Object.entries(extras).forEach(([key,short])=>{if(m[key])o[short]=m[key]});
  if(m.noAplica&&Object.values(m.noAplica).some(Boolean))o.na=m.noAplica;
  Object.keys(o).forEach(k=>{if(o[k]===""||o[k]==null)delete o[k]});return o}
// Master + piezas individuales: según la cantidad (entera, 1 a MAX_IND), sin importar la UM
const MAX_IND=100;
function esIndividual(m){return !!m&&Number.isInteger(m.cantidad)&&m.cantidad>=1&&m.cantidad<=MAX_IND}
// n = nº de pieza (sin n → QR master del registro completo)
function qrTexto(m,n){
  if(compacto())return `${BASE}#c=${encodeURIComponent(m.codigo)}${n?`&p=${n}-${m.cantidad}`:""}`;
  const o=payload(m);if(n){o.pn=n;o.pt=m.cantidad}else if((m.despachos||[]).length)o.x=existTxt(m)
  return `${BASE}#d=${b64e(JSON.stringify(o))}`}
// Devuelve {n,t} si el QR leído es de una pieza individual
function pzDe(txt){
  txt=String(txt||"");let x=/#(?:[^#]*&)?d=([^&\s]+)/.exec(txt);
  if(x){try{const o=JSON.parse(b64d(x[1]));return o.pn>0?{n:+o.pn,t:+o.pt||0}:null}catch{return null}}
  x=/#(?:[^#]*&)?p=(\d+)-(\d+)/.exec(txt);return x?{n:+x[1],t:+x[2]}:null}
function piezasSurt(m){const s=new Set();(m.despachos||[]).forEach(d=>(d.piezas||[]).forEach(n=>s.add(n)));return s}
// Extrae el código de un QR nuevo (enlace) o de uno antiguo (texto "Código: ...")
function codDe(txt){
  txt=String(txt||"");let x=/#(?:[^#]*&)?d=([^&\s]+)/.exec(txt);
  if(x){try{return String(JSON.parse(b64d(x[1])).c||"")}catch{return ""}}
  x=/#(?:[^#]*&)?c=([^&\s]+)/.exec(txt);
  if(x){try{return decodeURIComponent(x[1])}catch{return x[1]}}
  x=/C[óo]digo:\s*(.+)/i.exec(txt);return(x?x[1]:txt).split("\n")[0].trim()}
function qrURL(){const c=$("qrBox").querySelector("canvas");if(c)return c.toDataURL("image/png");const i=$("qrBox").querySelector("img");return i?i.src:""}

function dibujarQR(txt,px=240){
  let qr=null;
  for(const lvl of ["M","L"]){try{qr=qrcode(0,lvl);qr.addData(txt,"Byte");qr.make();break}catch(e){qr=null}}
  if(!qr)return null;
  const n=qr.getModuleCount(),cell=Math.max(4,Math.floor(px/n)),q=4,size=(n+2*q)*cell;
  const c=document.createElement("canvas");c.width=c.height=size;
  const x=c.getContext("2d");x.fillStyle="#fff";x.fillRect(0,0,size,size);x.fillStyle="#000";
  for(let r=0;r<n;r++)for(let k=0;k<n;k++)if(qr.isDark(r,k))x.fillRect((k+q)*cell,(r+q)*cell,cell,cell);
  c.style.cssText="width:220px;height:220px;image-rendering:pixelated";
  return c;
}
function mostrar(m){
  const pagina=document.querySelector('.section-tab[data-module="vReg"][aria-current="page"]')?.dataset.page;
  activarSubpantalla("vReg",pagina==="historial"?"historial":"captura");
  $("vReg").classList.add("qr-visible");
  actual=m;const box=$("qrBox");box.innerHTML="";
  const txt=qrTexto(m);
  const c=dibujarQR(txt);
  if(c)box.appendChild(c);else{box.textContent="No se pudo generar el QR: los datos son demasiado largos. Acorte descripción o ubicación.";actual=null}
  const rows=[["Código",m.codigo],["Almacén",m.ubicacionAlmacenNombre||almacenes.find(a=>a.codigo===m.ubicacionAlmacenCodigo)?.nombre],["Ubicación",m.ubicacion],["Clasificación",nombreTipoRegistro(m.tipoRegistro)],["Tipo material",m.tipoMaterial],["Descripción",m.descripcion],
    ["Sección / especificación",seccionEspecificacion(m)],["Longitud IND",m.longitudIndividual!=null?`${m.longitudIndividual} mm`:""],["Manufactura",m.manufactura],["N.º Heat",m.numeroHeat],["N.º lote MP",m.numeroLoteMP||m.lote],["N.º cert.",m.numeroCertificado],["Total mm",m.totalMM],["Fecha de recibo",m.fechaRecibo],["Cantidad",`${m.cantidad} ${m.um}`],
    ["Etiqueta",`MASTER · ${m.cantidad} ${m.um}`+(esIndividual(m)?` (+ ${m.cantidad} QR individuales)`:` (sin QR individuales: la cantidad debe ser entera y de 1 a ${MAX_IND})`)],
    ["Peso",m.peso!=null?`${m.peso}${m.pesoU||m.pesoUnidad?" "+(m.pesoU||m.pesoUnidad):m.pesoUnidadPendiente?" (unidad sin definir)":""}`:""],["Orden",m.ordenMaterial],["PO number",m.poNumber],
    ["Existencia",existTxt(m)],["Últ. modificación",m.fechaModificacion]];
  $("ficha").innerHTML=rows.filter(r=>r[1]).map(r=>`<div class="d"><b>${r[0]}</b><span>${esc(r[1])}</span></div>`).join("");
  $("bPrint").disabled=$("bPng").disabled=!actual;
  const ind=!!actual&&esIndividual(m);$("piezasBox").hidden=!ind;$("pzDet").open=false;$("pzGrid").innerHTML="";
  if(ind)$("piezasInfo").innerHTML=`<b>${m.cantidad} ${esc(m.um)}</b> (≤ ${MAX_IND}): además del QR master se generan <b>${m.cantidad} QR individuales</b>, uno por pieza (1/${m.cantidad} … ${m.cantidad}/${m.cantidad}).`;
}
function qrPieza(m,n){return dibujarQR(qrTexto(m,n))}
$("pzDet").addEventListener("toggle",()=>{
  const g=$("pzGrid");if(!$("pzDet").open||g.childElementCount||!actual||!esIndividual(actual))return;
  const m=actual;
  for(let n=1;n<=m.cantidad;n++){const c=qrPieza(m,n);if(!c)continue;c.style.cssText="width:96px;height:96px;image-rendering:pixelated";
    const d=document.createElement("div");d.className="pzi";d.dataset.n=n;d.style.cursor="pointer";d.title="Clic para descargar este PNG";d.appendChild(c);const t=document.createElement("div");t.textContent=`${n} / ${m.cantidad}`;d.appendChild(t);g.appendChild(d)}
});
// --- Impresión en Brother QL-800: 1 etiqueta por página; el papel se elige en el controlador ---
const ETQ={qr:50,lv:2,fs:11,cl:3};
function qrData(m,n){const c=dibujarQR(qrTexto(m,n),600);return c?c.toDataURL("image/png"):null}
function etqHTML(m,n,src){
  const f=ETQ,N=m.cantidad,tag=n?`PIEZA ${n} / ${N}`:`MASTER · ${esc(N)} ${esc(m.um)}`;
  let t=`<b class="cd">${esc(m.codigo)}</b><span class="tg">${tag}</span>`;
  if(f.lv>=1)t+=`<small class="ds">${esc(m.descripcion)}</small>`;
  if(f.lv>=2)t+=[m.lote?"Lote: "+m.lote:"",seccionEspecificacion(m),m.longitudIndividual!=null?`Long IND: ${m.longitudIndividual} mm`:"",m.peso!=null?`Peso: ${m.peso}${m.pesoU||m.pesoUnidad?" "+(m.pesoU||m.pesoUnidad):m.pesoUnidadPendiente?" (unidad sin definir)":""}`:"",m.ubicacion?"Ubic: "+m.ubicacion:"",m.fechaRecibo?"Recibo: "+m.fechaRecibo:""].filter(Boolean).map(x=>`<small>${esc(x)}</small>`).join("");
  return `<div class="pz${n?"":" mst"}"><img src="${src}" alt="QR"><div class="tx">${t}</div></div>`}
function cssEtq(){const f=ETQ;
  return `@media print{@page{size:auto;margin:0}
#etiqueta.hoja .pz{width:100%;height:100%;padding:1.5mm;gap:1.2mm;font-size:${f.fs}pt;flex-direction:column;align-items:center;justify-content:center;text-align:center}
#etiqueta.hoja .pz img{width:min(${f.qr}mm, 80vw);height:auto;aspect-ratio:1}
#etiqueta.hoja .pz .tx{width:100%;align-items:center}
#etiqueta.hoja .pz .ds{-webkit-line-clamp:${f.cl}}}`}
function limpiarEtqCss(){const st=$("etqCss");if(st)st.textContent=""}
function imprimirEtq(lista){
  const e=$("etiqueta");e.className="hoja";e.innerHTML=lista.map(x=>etqHTML(x.m,x.n,x.src)).join("");
  let st=$("etqCss");if(!st){st=document.createElement("style");st.id="etqCss";document.head.appendChild(st)}
  st.textContent=cssEtq();window.addEventListener("afterprint",limpiarEtqCss,{once:true});
  Promise.all([...e.querySelectorAll("img")].map(i=>i.decode?i.decode().catch(()=>{}):0)).then(()=>window.print());
}
function imprimirPiezas(conMaster){
  if(!actual||!esIndividual(actual))return;const m=actual,L=[];
  if(conMaster){const u=qrData(m,0);if(!u)return;L.push({m,n:0,src:u})}
  for(let n=1;n<=m.cantidad;n++){const u=qrData(m,n);if(!u){toast("No se pudo generar el QR de la pieza "+n,1);return}L.push({m,n,src:u})}
  imprimirEtq(L);
}
$("bPrintPz").onclick=()=>imprimirPiezas(false);
// --- Descarga de PNG: ZIP (sin compresión) con un PNG por pieza ---
const CRCT=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc32=b=>{let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRCT[(c^b[i])&255]^(c>>>8);return(c^0xFFFFFFFF)>>>0};
function zipBytes(files){ // files: [{name,data:Uint8Array}]
  const enc=new TextEncoder(),now=new Date(),tm=(now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1),dt=((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
  const parts=[],cen=[];let off=0;
  files.forEach(f=>{
    const nm=enc.encode(f.name),crc=crc32(f.data),sz=f.data.length;
    const lh=new DataView(new ArrayBuffer(30));lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);lh.setUint16(8,0,true);lh.setUint16(10,tm,true);lh.setUint16(12,dt,true);lh.setUint32(14,crc,true);lh.setUint32(18,sz,true);lh.setUint32(22,sz,true);lh.setUint16(26,nm.length,true);lh.setUint16(28,0,true);
    parts.push(new Uint8Array(lh.buffer),nm,f.data);
    const ch=new DataView(new ArrayBuffer(46));ch.setUint32(0,0x02014b50,true);ch.setUint16(4,20,true);ch.setUint16(6,20,true);ch.setUint16(8,0x0800,true);ch.setUint16(10,0,true);ch.setUint16(12,tm,true);ch.setUint16(14,dt,true);ch.setUint32(16,crc,true);ch.setUint32(20,sz,true);ch.setUint32(24,sz,true);ch.setUint16(28,nm.length,true);ch.setUint32(42,off,true);
    cen.push(new Uint8Array(ch.buffer),nm);off+=30+nm.length+sz});
  const cs=cen.reduce((a,b)=>a+b.length,0),end=new DataView(new ArrayBuffer(22));
  end.setUint32(0,0x06054b50,true);end.setUint16(8,files.length,true);end.setUint16(10,files.length,true);end.setUint32(12,cs,true);end.setUint32(16,off,true);
  return new Blob([...parts,...cen,new Uint8Array(end.buffer)],{type:"application/zip"})}
const durl=u=>Uint8Array.from(atob(u.split(",")[1]),c=>c.charCodeAt(0));
const safeN=t=>String(t).replace(/[^\w.\-]+/g,"_");
const nomPz=(m,n)=>`QR_${safeN(m.codigo)}_pieza_${String(n).padStart(String(m.cantidad).length,"0")}_de_${m.cantidad}.png`;
async function bajarBlob(nombre,blob){
  if(matchMedia("(pointer:coarse)").matches){try{const f=new File([blob],nombre,{type:blob.type});if(navigator.canShare&&navigator.canShare({files:[f]})){await navigator.share({files:[f],title:nombre});return}}catch(e){if(e.name==="AbortError")return}}
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=nombre;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
$("bPngPz").onclick=()=>{
  if(!actual||!esIndividual(actual))return;const m=actual,F=[];
  for(let n=1;n<=m.cantidad;n++){const u=qrData(m,n);if(!u){toast("No se pudo generar el QR de la pieza "+n,1);return}F.push({name:nomPz(m,n),data:durl(u)})}
  bajarBlob(`QR_${safeN(m.codigo)}_piezas.zip`,zipBytes(F));toast(`${F.length} PNG en ZIP ⬇️`);
};
$("pzGrid").addEventListener("click",e=>{
  const d=e.target.closest(".pzi");if(!d||!actual)return;const n=+d.dataset.n,u=qrData(actual,n);if(!u)return;
  bajarBlob(nomPz(actual,n),new Blob([durl(u)],{type:"image/png"}));
});
$("bPrintAll").onclick=()=>imprimirPiezas(true);
$("bPng").onclick=()=>{const u=qrURL();if(!u||!actual)return;const a=document.createElement("a");a.href=u;a.download=`QR_${actual.codigo}.png`;a.click()};
$("bPrint").onclick=()=>{if(!actual)return;const u=qrData(actual,0);if(!u){toast("No se pudo generar el QR",1);return}imprimirEtq([{m:actual,n:0,src:u}])};

// --- Tabla ---
function render(){
  const codigos=new Set(mats.map(m=>m.codigo));seleccionados.forEach(c=>{if(!codigos.has(c))seleccionados.delete(c)});
  const q=$("buscar").value.toLowerCase().trim(),t=$("ftipo").value;
  let lista=mats.map((m,i)=>({m,i})).filter(({m})=>(!t||(t==="mp"?m.esMP:t==="consumibles"?m.tipoRegistro==="CONSUMIBLES":m.tipoRegistro==="GENERAL"))&&(!q||[m.codigo,m.descripcion,m.lote,m.proveedor,m.ubicacion,m.tipoMaterial,m.especificacionMedida,m.longitudIndividual,m.manufactura,m.numeroHeat,m.numeroLoteMP,m.numeroCertificado,m.totalMM,m.tipoMoneda,m.ordenMaterial,m.poNumber].join(" ").toLowerCase().includes(q)));
  if(orden.k)lista.sort((a,b)=>{const x=a.m[orden.k],y=b.m[orden.k];const r=typeof x==="number"?x-y:String(x??"").localeCompare(String(y??""),"es",{numeric:true});return orden.asc?r:-r});
  $("tbody").innerHTML=lista.map(({m,i})=>`<tr>
    <td>${esAdmin?`<input type="checkbox" class="row-select" data-select="${esc(m.codigo)}" aria-label="Seleccionar ${esc(m.codigo)}" ${seleccionados.has(m.codigo)?"checked":""}>`:""}</td>
    <td><b>${esc(m.codigo)}</b></td><td class="l">${esc(m.descripcion)}</td><td>${esc(m.cantidad)} PZ · ${esc(existTxt(m))}</td>
    <td>${m.esMP?'<span class="tag t1">M. PRIMA</span>':m.tipoRegistro==="CONSUMIBLES"?'<span class="tag t2">CONSUMIBLES</span>':'<span class="tag t2">GENERAL</span>'}</td>
    <td>${esc(m.ubicacion||"Sin ubicación")}</td><td>${esc(m.fechaRecibo)}</td>
    <td style="white-space:nowrap"><button type="button" class="ic" data-a="qr" data-i="${i}" title="Ver QR">🔳</button>${esAdmin?`<button type="button" class="ic" data-a="ed" data-i="${i}" title="Editar">✏️</button><button type="button" class="ic" data-a="del" data-i="${i}" title="Eliminar">🗑️</button>`:""}</td></tr>`).join("");
  $("vacio").hidden=lista.length>0;
  const mp=mats.filter(m=>m.esMP).length,consumibles=mats.filter(m=>m.tipoRegistro==="CONSUMIBLES").length;
  $("sTot").textContent=`${mats.length} registros`;$("sMp").textContent=`${mp} materia prima`;$("sConsumibles").textContent=`${consumibles} consumibles`;$("sGen").textContent=`${mats.length-mp-consumibles} generales`;
  syncBulkActions();
}
function renderInventario(){
  const selector=$("invAlmacen"),almacenActual=selector.value;
  selector.innerHTML='<option value="">Todos los almacenes</option>'+almacenes.map(a=>`<option value="${esc(a.codigo)}">${esc(a.nombre)}</option>`).join("");
  selector.value=almacenes.some(a=>a.codigo===almacenActual)?almacenActual:"";
  const q=$("invBuscar").value.trim().toLocaleLowerCase(),tipo=$("invTipo").value,almacenFiltro=selector.value;
  const ubiLoc=ubiFiltroCodigo?ubicaciones.find(loc=>loc.codigo===ubiFiltroCodigo):null;
  const permitidos=ubiLoc?new Set(listaMaterialesEnUbicacion(ubiLoc).map(m=>m.codigo)):null;
  const lista=mats.filter(m=>{const loc=ubiPorCodigo(m.ubicacionCodigo),almacenCodigo=loc?.almacen_codigo||m.ubicacionAlmacenCodigo,alm=almPorCodigo(almacenCodigo);return(!permitidos||permitidos.has(m.codigo))&&(!tipo||(tipo==="mp"?m.esMP:tipo==="consumibles"?m.tipoRegistro==="CONSUMIBLES":m.tipoRegistro==="GENERAL"))&&(!almacenFiltro||almacenCodigo===almacenFiltro)&&(!q||[m.codigo,m.descripcion,m.ubicacion,m.ubicacionCodigo,loc?.nombre,alm?.nombre,m.lote,m.numeroHeat,m.numeroLoteMP].join(" ").toLocaleLowerCase().includes(q))});
  $("invBody").innerHTML=lista.map(m=>{const loc=ubiPorCodigo(m.ubicacionCodigo),alm=almPorCodigo(loc?.almacen_codigo||m.ubicacionAlmacenCodigo),tag=m.esMP?'<span class="tag t1">M. PRIMA</span>':m.tipoRegistro==="CONSUMIBLES"?'<span class="tag t2">CONSUMIBLES</span>':'<span class="tag t2">GENERAL</span>';return`<tr><td><b>${esc(m.codigo)}</b><br><small>${tag}</small></td><td class="l">${esc(m.descripcion)}</td><td>${esc(existTxt(m))}</td><td class="l">${esc(alm?.nombre||m.ubicacionAlmacenNombre||"Sin almacén")} › ${esc(loc?ubicacionConCodigo(loc):m.ubicacion||"Sin ubicación")}</td></tr>`}).join("");
  $("invVacio").hidden=lista.length>0;
  actualizarEstadoConteoCiclico();
}
function actualizarEstadoConteoCiclico(){
  const estado=$("conteoEstado");
  estado.hidden=!ajustesInventarioError;
  estado.textContent=ajustesInventarioError;
}
$("invBuscar").oninput=deb(renderInventario);$("invTipo").onchange=$("invAlmacen").onchange=renderInventario;
let ubiConsultaCodigo=null,ubiFiltroCodigo=null,conteoTrasladadoPendiente=false;
function limpiarConsultaUbicacion(){
  ubiConsultaCodigo=null;
  ubiFiltroCodigo=null;
  $("ubiConsultaBox").hidden=true;
  $("bUbiConsultaLimpiar").hidden=true;
  $("ubiConsultaInfo").textContent="-";
  $("ubiConsultaCnt").textContent="0 materiales";
  $("ubiConsultaBody").innerHTML="";
  $("ubiConsultaVacio").hidden=true;
  $("ubiConsultaBuscar").value="";
  $("ubiConsultaBusqueda").hidden=true;
  $("ubiConsultaResultados").innerHTML="";
  renderInventario();
}
function ubicacionPorCodigo(codigo){
  const clave=String(codigo||"").trim().toLocaleLowerCase();
  const canonica=normalizarCodigoUbicacion(codigo).toLocaleLowerCase();
  const directa=ubicaciones.find(loc=>loc.codigo.toLocaleLowerCase()===clave);
  if(directa)return directa;
  const alias=ubicacionesAlias.find(item=>item.codigo_alias.toLocaleLowerCase()===clave);
  const destinoAlias=alias&&ubicaciones.find(loc=>loc.codigo===alias.ubicacion_codigo);
  if(destinoAlias)return destinoAlias;
  const directaCanonica=ubicaciones.find(loc=>normalizarCodigoUbicacion(loc.codigo).toLocaleLowerCase()===canonica);
  if(directaCanonica)return directaCanonica;
  const aliasCanonico=ubicacionesAlias.find(item=>normalizarCodigoUbicacion(item.codigo_alias).toLocaleLowerCase()===canonica);
  const destinoCanonico=aliasCanonico&&ubicaciones.find(loc=>loc.codigo===aliasCanonico.ubicacion_codigo);
  if(destinoCanonico)return destinoCanonico;
  return ubicaciones.find(loc=>String(loc.nombre||"").trim().toLocaleLowerCase()===clave)||null;
}
function mostrarMaterialesPorUbicacion(codigo){
  const loc=ubicacionPorCodigo(codigo);
  if(!loc){toast(ubicacionesError||ubicacionesAliasError||`Ubicación no registrada: ${codigo}`,1);return}
  const almacen=almacenes.find(a=>a.codigo===loc.almacen_codigo);
  const lista=listaMaterialesEnUbicacion(loc);
  ubiConsultaCodigo=loc.codigo;
  ubiFiltroCodigo=loc.codigo;
  $("ubiConsultaInfo").innerHTML=`${esc(almacen?.nombre||"Sin almacén")} · <b>${esc(ubicacionConCodigo(loc))}</b>`;
  $("ubiConsultaCnt").textContent=`${lista.length} material${lista.length===1?"":"es"}`;
  actualizarEstadoConteoCiclico();
  $("ubiConsultaBusqueda").hidden=!esAdmin;
  $("ubiConsultaBuscar").value="";
  $("ubiConsultaResultados").innerHTML='<p class="ubicaciones-ayuda">Busca en el catálogo del módulo Registro. Para ajustar existencias, el material debe estar asignado a esta ubicación.</p>';
  $("ubiConsultaBody").innerHTML=lista.map(m=>{const almM=almacenes.find(a=>a.codigo===(ubiPorCodigo(m.ubicacionCodigo)?.almacen_codigo||m.ubicacionAlmacenCodigo));return`<tr><td><b>${esc(m.codigo)}</b></td><td class="l">${esc(m.descripcion||"")}</td><td>${m.esMP?'<span class="tag t1">M. PRIMA</span>':m.tipoRegistro==="CONSUMIBLES"?'<span class="tag t2">CONSUMIBLES</span>':'<span class="tag t2">GENERAL</span>'}</td><td>${esc(existTxt(m))}</td><td>${esc(almM?.nombre||m.ubicacionAlmacenNombre||almacen?.nombre||"Sin almacén")}</td><td>${esAdmin?`<button type="button" class="bs" data-conteo="${esc(m.codigo)}" ${ajustesInventarioError?"disabled title=\"No se pudo cargar el historial de conteos\"":""}>📋 Conteo cíclico</button>`:"—"}</td></tr>`}).join("");
  $("ubiConsultaBody").querySelectorAll("[data-conteo]").forEach(btn=>btn.onclick=()=>abrirConteoCiclico(btn.dataset.conteo,loc));
  $("ubiConsultaVacio").hidden=lista.length>0;
  $("ubiConsultaBox").hidden=false;
  $("bUbiConsultaLimpiar").hidden=false;
  document.querySelector('.tab[data-v="vInv"]').click();
  const tab=document.querySelector('.section-tab[data-module="vInv"][data-page="existencias"]');if(tab)tab.click();
  $("ubiConsultaBox").scrollIntoView({behavior:"smooth",block:"nearest"});
  renderInventario();
}
$("bCamUbiConsulta").onclick=()=>{modoEscaneo="ubiConsulta";camOn()};
$("bUbiConsultaLimpiar").onclick=()=>{limpiarConsultaUbicacion()};
function buscarMaterialesParaConteo(){
  const resultados=$("ubiConsultaResultados"),consulta=$("ubiConsultaBuscar").value.trim().toLocaleLowerCase();
  if(!consulta){
    resultados.innerHTML='<p class="ubicaciones-ayuda">Busca en el catálogo del módulo Registro por código o descripción.</p>';
    return;
  }
  const coincidencias=mats.filter(material=>[material.codigo,material.descripcion].some(valor=>String(valor||"").toLocaleLowerCase().includes(consulta)));
  if(!coincidencias.length){
    resultados.innerHTML='<p class="ubicaciones-ayuda">No se encontró ningún material registrado con ese código o descripción.</p>';
    return;
  }
  const ubicacion=ubicaciones.find(loc=>loc.codigo===ubiConsultaCodigo);
  if(!ubicacion){resultados.innerHTML='<p class="ubicaciones-ayuda">Escanea de nuevo una ubicación para consultar los materiales.</p>';return}
  const visibles=coincidencias.slice(0,20);
  resultados.innerHTML=visibles.map(material=>{
    const enUbicacion=listaMaterialesEnUbicacion(ubicacion).some(item=>item.codigo===material.codigo);
    const ubicacionActual=ubicaciones.find(loc=>loc.codigo===material.ubicacionCodigo)
      ||ubicaciones.find(loc=>!material.ubicacionCodigo&&String(material.ubicacion||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase()&&(!material.ubicacionAlmacenCodigo||material.ubicacionAlmacenCodigo===loc.almacen_codigo));
    const ubicacionTexto=ubicacionActual?ubicacionConCodigo(ubicacionActual):material.ubicacion||"Sin ubicación";
    const accion=enUbicacion?"📋 Conteo cíclico":ubicacionActual||material.ubicacion?"Mover y contar":"Agregar y contar";
    return`<div class="ubicacion-material"><span><b>${esc(material.codigo)}</b><span>${esc(material.descripcion||"Sin descripción")} · ${esc(existTxt(material))}<br><small>${enUbicacion?"Asignado a esta ubicación":`Ubicación actual: ${esc(ubicacionTexto)}`}</small></span></span><button type="button" class="bs" data-conteo-busqueda="${esc(material.codigo)}" ${ajustesInventarioError?"disabled title=\"No se pudo cargar el historial de conteos\"":""}>${accion}</button></div>`;
  }).join("")+(coincidencias.length>visibles.length?`<p class="ubicaciones-ayuda">Se muestran los primeros ${visibles.length} resultados de ${coincidencias.length}. Afina la búsqueda para encontrar el material.</p>`:"");
  resultados.querySelectorAll("[data-conteo-busqueda]").forEach(btn=>btn.onclick=()=>agregarMoverMaterialParaConteo(btn.dataset.conteoBusqueda,ubicacion,btn));
}
$("ubiConsultaBuscar").addEventListener("input",deb(buscarMaterialesParaConteo));
async function agregarMoverMaterialParaConteo(codigo,ubicacion,boton){
  if(!esAdmin){toast("Solo el administrador puede agregar materiales o ajustar existencias.",1);return}
  if(ajustesInventarioError){toast(ajustesInventarioError,1);return}
  const material=mats.find(item=>item.codigo===codigo);
  if(!material){toast("No se encontró el material. Actualiza el inventario e inténtalo de nuevo.",1);return}
  if(listaMaterialesEnUbicacion(ubicacion).some(item=>item.codigo===codigo)){
    abrirConteoCiclico(codigo,ubicacion);
    return;
  }
  const anterior=ubicaciones.find(loc=>loc.codigo===material.ubicacionCodigo)
    ||ubicaciones.find(loc=>!material.ubicacionCodigo&&String(material.ubicacion||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase()&&(!material.ubicacionAlmacenCodigo||material.ubicacionAlmacenCodigo===loc.almacen_codigo));
  const anteriorTexto=anterior?ubicacionConCodigo(anterior):material.ubicacion||"Sin ubicación";
  const accion=anterior?"mover":"agregar";
  if(!await confirmar(`¿${anterior?"Mover":"Agregar"} ${material.codigo} a ${ubicacionConCodigo(ubicacion)} y después hacer el conteo manual?\n\nUbicación actual: ${anteriorTexto}\nExistencia disponible actual: ${existTxt(material)}\n\n${anterior?"Se registrará el traspaso en el historial.":"El material quedará asignado a esta ubicación."}`))return;
  const actual=existencia(material);
  const folio=`CONTEO-${Date.now()}`;
  const payload=[{codigo:material.codigo,existencia:{modo:actual.modo,cantidad:actual.modo==="pz"?actual.rest:actual.qty,longitud:actual.modo==="long"?actual.rest:null,unidad:actual.u,peso:actual.peso?.peso??null,pesoUnidad:actual.peso?.u??null}}];
  boton.disabled=true;
  try{
    const{data,error}=await sb.rpc("asignar_inventario_ubicacion",{p_ubicacion_codigo:ubicacion.codigo,p_folio:folio,p_items:payload});
    if(error)throw error;
    if(Number(data)!==1)throw new Error(`Respuesta incompleta del servidor (${data??"sin cantidad"})`);
  }catch(err){
    errDB(err,`No se pudo ${accion} el material a la ubicación`);
    boton.disabled=false;
    return;
  }
  try{
    await Promise.all([cargar(),cargarUbicaciones()]);
  }catch(err){
    errDB(err,`${material.codigo} ya se agregó a ${ubicacionConCodigo(ubicacion)}, pero no se pudo actualizar la vista`);
    return;
  }
  const consulta=$("ubiConsultaBuscar").value;
  mostrarMaterialesPorUbicacion(ubicacion.codigo);
  $("ubiConsultaBuscar").value=consulta;
  buscarMaterialesParaConteo();
  if(abrirConteoCiclico(material.codigo,ubicacion))conteoTrasladadoPendiente=true;
}
function abrirConteoCiclico(codigo,loc){
  if(!esAdmin){toast("Solo el administrador puede ajustar existencias.",1);return false}
  if(ajustesInventarioError){toast(ajustesInventarioError,1);return false}
  const material=mats.find(m=>m.codigo===codigo);
  if(!material||!listaMaterialesEnUbicacion(loc).some(m=>m.codigo===codigo)){toast("El material ya no está asignado a esta ubicación. Actualiza la consulta.",1);return false}
  const e=existencia(material);
  $("conteoCodigo").value=material.codigo;
  $("conteoUbicacion").value=loc.codigo;
  $("conteoAnterior").value=String(e.rest);
  $("conteoMaterial").textContent=`${material.codigo} · ${material.descripcion||""} · ${ubicacionConCodigo(loc)}`;
  $("conteoActual").textContent=existTxt(material);
  $("conteoUnidad").textContent=e.u;
  $("conteoCantidad").value=String(Math.max(0,e.rest));
  $("conteoMotivo").value="";
  conteoTrasladadoPendiente=false;
  $("conteoDialog").showModal();
  $("conteoCantidad").focus();
  return true;
}
$("conteoCancelar").onclick=()=>$("conteoDialog").close();
$("conteoForm").addEventListener("submit",async event=>{
  event.preventDefault();
  if(!esAdmin){toast("Solo el administrador puede ajustar existencias.",1);$("conteoDialog").close();return}
  const codigo=$("conteoCodigo").value,ubicacionCodigo=$("conteoUbicacion").value,anterior=Number($("conteoAnterior").value),contada=r4(Number($("conteoCantidad").value));
  const material=mats.find(m=>m.codigo===codigo),ubicacion=ubicaciones.find(loc=>loc.codigo===ubicacionCodigo);
  if(!material||!ubicacion||!listaMaterialesEnUbicacion(ubicacion).some(m=>m.codigo===codigo)){toast("El material cambió de ubicación. Escanea la ubicación de nuevo.",1);$("conteoDialog").close();return}
  if(!Number.isFinite(contada)||contada<0){toast("Ingresa una existencia contada válida.",1);return}
  const actualCalculado=existencia(material);
  if(Math.abs(actualCalculado.rest-anterior)>0.0001){toast("La existencia cambió mientras realizabas el conteo. Vuelve a abrir el ajuste.",1);$("conteoDialog").close();return}
  $("conteoGuardar").disabled=true;
  try{
    const{error}=await sb.rpc("registrar_conteo_ciclico",{p_codigo_material:codigo,p_ubicacion_codigo:ubicacionCodigo,p_existencia_anterior:anterior,p_existencia_contada:contada,p_unidad:actualCalculado.u,p_motivo:$("conteoMotivo").value.trim()});
    if(error)throw error;
    $("conteoDialog").close();
    conteoTrasladadoPendiente=false;
    try{await cargar()}catch(err){errDB(err,"Conteo guardado, pero no se pudo actualizar el inventario");return}
    renderHistUbicaciones();
    if(ubiConsultaCodigo)mostrarMaterialesPorUbicacion(ubiConsultaCodigo);
    toast(`Conteo guardado para ${codigo}. Ajuste: ${fmt(contada-anterior)} ${actualCalculado.u}`);
  }catch(err){errDB(err,conteoTrasladadoPendiente?"El material ya se agregó a la ubicación, pero no se pudo guardar el conteo cíclico":"No se pudo guardar el conteo cíclico")}
  finally{$("conteoGuardar").disabled=false}
});
function procesarEscaneoUbiConsulta(txt){
  const raw=String(txt||"").replace(/^\uFEFF/,"").trim(),prefijo="TRAZABILIDAD-UBICACION:";
  if(raw.toUpperCase().startsWith(prefijo)){modoEscaneo="inventario";mostrarMaterialesPorUbicacion(raw.slice(prefijo.length).trim());camOff();return true}
  const loc=ubicacionPorCodigo(raw);if(loc){modoEscaneo="inventario";mostrarMaterialesPorUbicacion(loc.codigo);camOff();return true}
  toast("Escanee un QR de ubicación válido",1);return false;
}
function renderAlmacenes(){
  const estado=$("almacenEstado"),lista=$("almacenLista"),form=$("almacenForm");
  if(almacenesError){estado.innerHTML=`<div class="res bad">${esc(almacenesError)} <a href="./supabase_almacenes.sql" download>Descargar migración</a></div>`;form.hidden=true;lista.innerHTML="";return}
  estado.innerHTML=[ubicacionesError,movimientosError].filter(Boolean).map(error=>`<div class="res bad">${esc(error)} <a href="./supabase_almacenes.sql" download>Ver migración</a></div>`).join("");form.hidden=false;
  if(!almacenes.length){lista.innerHTML='<div class="empty">No hay almacenes registrados.</div>';return}
  const abiertas=new Set([...lista.querySelectorAll(".ubicacion-card[open]")].map(item=>item.dataset.codigo));
  lista.innerHTML=almacenes.map(a=>{
    const locs=ubicaciones.filter(loc=>loc.almacen_codigo===a.codigo),tipo={MP:"Materia prima",GENERAL:"Material general",CONSUMIBLES:"Consumibles",OTRO:"Otro"}[a.clasificacion]||"Sin clasificación";
    const materiales=materialesEnAlmacenes([a.codigo]),movimientos=movimientosEnAlmacenes([a.codigo],locs);
    return`<details class="ubicacion-card" data-codigo="${esc(a.codigo)}" ${abiertas.has(a.codigo)?"open":""}><summary><span class="ubicacion-nombre">${esc(a.nombre)}</span><span class="ubicacion-codigo">${esc(a.codigo)} · ${locs.length} ubicaciones</span></summary><div class="ubicacion-detalle"><div class="d"><b>Clasificación</b><span>${esc(tipo)}</span></div><div class="d"><b>Ubicaciones</b><span>${locs.length?locs.map(loc=>esc(ubicacionConCodigo(loc))).join(" · "):"Sin ubicaciones"}</span></div>${esAdmin?`<div class="res bad">Al eliminar se conservarán los materiales, pero se quitará su almacén y ubicación y se borrarán ${movimientos} movimientos del historial${materiales?` (${materiales} material(es) quedarán sin ubicación)`: ""}.</div><div class="ubicacion-admin-actions"><button type="button" class="bd" data-eliminar-almacen="${esc(a.codigo)}">🗑️ Eliminar almacén y su historial</button></div>`:""}</div></details>`
  }).join("");
  lista.querySelectorAll("[data-eliminar-almacen]").forEach(btn=>btn.onclick=()=>eliminarAlmacen(btn.dataset.eliminarAlmacen));
}
function listaMaterialesEnUbicacion(loc){
  return mats.filter(m=>m.ubicacionCodigo===loc.codigo||(!m.ubicacionCodigo&&String(m.ubicacion||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase()&&(!m.ubicacionAlmacenCodigo||m.ubicacionAlmacenCodigo===loc.almacen_codigo)));
}
function materialesEnUbicacion(loc){return listaMaterialesEnUbicacion(loc).length}
function movimientosDeUbicacion(loc){
  return movimientosUbicacion.filter(m=>
    m.ubicacion_anterior_codigo===loc.codigo||m.ubicacion_nueva_codigo===loc.codigo||
    (!m.ubicacion_anterior_codigo&&String(m.ubicacion_anterior||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase())||
    (!m.ubicacion_nueva_codigo&&String(m.ubicacion_nueva||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase())
  ).length;
}
function materialesEnAlmacenes(codigos){
  const codSet=new Set(codigos),locs=ubicaciones.filter(loc=>codSet.has(loc.almacen_codigo));
  return mats.filter(m=>codSet.has(m.ubicacionAlmacenCodigo)||locs.some(loc=>(
    m.ubicacionCodigo===loc.codigo||(!m.ubicacionCodigo&&String(m.ubicacion||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase())
  ))).length;
}
function movimientosEnAlmacenes(codigos,locs){
  const codSet=new Set(codigos),locSet=new Set(locs.map(loc=>loc.codigo));
  return movimientosUbicacion.filter(m=>
    codSet.has(m.ubicacion_anterior_almacen_codigo)||codSet.has(m.ubicacion_nueva_almacen_codigo)||
    locSet.has(m.ubicacion_anterior_codigo)||locSet.has(m.ubicacion_nueva_codigo)||
    locs.some(loc=>
      (!m.ubicacion_anterior_codigo&&String(m.ubicacion_anterior||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase())||
      (!m.ubicacion_nueva_codigo&&String(m.ubicacion_nueva||"").trim().toLocaleLowerCase()===loc.nombre.toLocaleLowerCase())
    )
  ).length;
}
async function eliminarUbicacion(codigo){
  if(!esAdmin){toast("Solo el administrador puede eliminar ubicaciones.",1);return}
  if(movimientosError){toast("No se pudo verificar el historial. Recarga antes de eliminar esta ubicación.",1);return}
  const loc=ubicaciones.find(item=>item.codigo===codigo);if(!loc)return;
  const materiales=materialesEnUbicacion(loc),movimientos=movimientosDeUbicacion(loc);
  if(materiales){toast("Traspasa primero todo el inventario de esta ubicación.",1);return}
  if(!await confirmar(`¿Eliminar la ubicación ${ubicacionConCodigo(loc)} del almacén ${almacenes.find(a=>a.codigo===loc.almacen_codigo)?.nombre||"Sin almacén"}?${movimientos?`\n\nSe conservarán ${movimientos} movimiento(s) en el historial con el nombre y código originales.`:""}`))return;
  let movimientosConservados=movimientos;
  try{
    const{data,error}=await sb.rpc("eliminar_ubicacion_vacia",{p_ubicacion_codigo:codigo});if(error)throw error;
    movimientosConservados=Number(data?.movimientos_conservados??movimientos);
  }catch(err){
    if(/UBICACION_NO_VACIA/.test(String(err?.message||err)))toast("La ubicación todavía tiene inventario. Traspásalo antes de eliminarla.",1);
    else errDB(err,"No se pudo eliminar la ubicación");
    return;
  }
  await cargarUbicaciones();toast(`Ubicación eliminada: ${ubicacionConCodigo(loc)}${movimientosConservados?` · ${movimientosConservados} movimiento(s) conservados`:""}`);
}
async function eliminarAlmacen(codigo){
  if(!esAdmin){toast("Solo el administrador puede eliminar almacenes.",1);return}
  const almacen=almacenes.find(item=>item.codigo===codigo);if(!almacen)return;
  const locs=ubicaciones.filter(loc=>loc.almacen_codigo===codigo),materiales=materialesEnAlmacenes([codigo]),movimientos=movimientosEnAlmacenes([codigo],locs);
  const detalle=`\n\nSe eliminarán ${locs.length} ubicación(es) y ${movimientos} movimientos del historial. Los ${materiales} materiales se conservarán, pero quedarán sin almacén ni ubicación. Esta acción no se puede deshacer.`;
  if(!await confirmar({titulo:"Eliminar almacén",mensaje:`¿Eliminar el almacén ${almacen.nombre} (${almacen.codigo})?${detalle}`,escribir:almacen.codigo,boton:"Eliminar almacén",peligro:true}))return;
  try{
    const{data,error}=await sb.rpc("eliminar_almacen_inventario",{p_almacen_codigo:codigo});if(error)throw error;
    const borrados=Number(data?.movimientos_eliminados??movimientos),liberados=Number(data?.materiales_desvinculados??materiales);
    await cargarUbicaciones();
    toast(`Almacén eliminado. ${liberados} materiales conservados sin ubicación · ${borrados} movimientos eliminados.`);
  }catch(err){errDB(err,"No se pudo eliminar el almacén")}
}
function actualizarSelectorAlmacenes(seleccionActual){
  const selector=$("ubiAlmacen"),valor=seleccionActual||selector.value;
  selector.innerHTML='<option value="">Seleccione un almacén…</option>'+almacenes.map(a=>`<option value="${esc(a.codigo)}">${esc(a.nombre)}${a.clasificacion?` · ${esc(a.clasificacion)}`:""}</option>`).join("");
  const predeterminado=almacenes.some(a=>a.codigo==="ALM-MP")?"ALM-MP":almacenes[0]?.codigo||"";
  selector.value=almacenes.some(a=>a.codigo===valor)?valor:predeterminado;
  $("ubiGuardar").disabled=!selector.value||!!almacenesError||!!ubicacionesError;
}
function actualizarDatalistUbicaciones(){
  const dl=$("ubicacionesDatalist");if(!dl)return;
  const opciones=[...new Map(ubicaciones.map(loc=>{
    const label=ubicacionConCodigo(loc);
    return [label.toLowerCase(), label];
  })).values()].sort((a,b)=>a.localeCompare(b,"es"));
  dl.innerHTML=opciones.map(v=>`<option value="${esc(v)}"></option>`).join("");
}
function normalizarCodigoUbicacion(codigo){
  const valor=String(codigo||"").trim();
  return /^UBI-/i.test(valor)?`UBI-${valor.slice(4)}`:`UBI-${valor}`;
}
const nombreUbicacion=loc=>{if(!loc)return"";const nombre=String(loc.nombre||"").trim();return nombre&&normalizarCodigoUbicacion(nombre).toLocaleLowerCase()!==normalizarCodigoUbicacion(loc.codigo).toLocaleLowerCase()?nombre:normalizarCodigoUbicacion(loc.codigo)};
const ubicacionConCodigo=loc=>{if(!loc)return"";const codigo=normalizarCodigoUbicacion(loc.codigo),n=nombreUbicacion(loc);return n===codigo?codigo:`${n} (${codigo})`};
function renderUbicaciones(){
  const estado=$("ubiEstado"),lista=$("ubiLista"),form=$("ubiForm");
  actualizarSelectorAlmacenes();
  if(almacenesError||ubicacionesError){estado.innerHTML=[almacenesError,ubicacionesError].filter(Boolean).map(error=>`<div class="res bad">${esc(error)} <a href="./supabase_almacenes.sql" download>Descargar migración</a></div>`).join("");form.hidden=true;lista.innerHTML="";return}
  estado.innerHTML=[movimientosError,ubicacionesAliasError?`${ubicacionesAliasError} Los QR nuevos siguen disponibles; los códigos impresos anteriores podrían no reconocerse.`:""].filter(Boolean).map(error=>`<div class="res bad">${esc(error)}</div>`).join("");form.hidden=false;
  const almacenCodigo=$("ubiAlmacen").value;
  if(!almacenCodigo){lista.innerHTML='<div class="empty">Crea o selecciona un almacén antes de agregar ubicaciones.</div>';return}
  const ubicacionesAlmacen=ubicaciones.filter(loc=>loc.almacen_codigo===almacenCodigo);
  if(!ubicacionesAlmacen.length){lista.className="ubicaciones-list";lista.innerHTML='<div class="empty">Este almacén todavía no tiene ubicaciones.</div>';return}
  const abiertas=new Set([...lista.querySelectorAll(".ubicacion-card[open]")].map(item=>item.dataset.codigo));
  lista.className="ubicaciones-list";
  lista.innerHTML=ubicacionesAlmacen.map(loc=>{
    const i=ubicaciones.indexOf(loc),alm=almacenes.find(a=>a.codigo===loc.almacen_codigo),materiales=listaMaterialesEnUbicacion(loc),existencias=materiales.length,movimientos=movimientosDeUbicacion(loc),ocupada=existencias>0||!!movimientosError;
    const materialesHtml=materiales.length
      ?`<section class="ubicacion-materiales" aria-label="Materiales en esta ubicación"><h3>Materiales en esta ubicación</h3><div class="ubicacion-material-list">${materiales.map(m=>`<div class="ubicacion-material"><span><b>${esc(m.codigo)}</b><span>${esc(m.descripcion)}</span></span><strong>${esc(existTxt(m))}</strong></div>`).join("")}</div></section>`
      :'<p class="ubicaciones-ayuda">No hay materiales en esta ubicación.</p>';
    const codigoVisible=normalizarCodigoUbicacion(loc.codigo),nombreVisible=nombreUbicacion(loc);
    return`<details class="ubicacion-card ubicacion-fila" data-index="${i}" data-codigo="${esc(loc.codigo)}" ${abiertas.has(loc.codigo)?"open":""}><summary><span class="ubicacion-nombre">${esc(codigoVisible)}</span><span class="ubicacion-codigo">${esc(alm?.nombre||"Sin almacén")}${nombreVisible!==codigoVisible?` · ${esc(nombreVisible)}`:""}</span></summary><div class="ubicacion-detalle"><div class="d"><b>Almacén</b><span>${esc(alm?.nombre||"Sin almacén")}</span></div><div class="d"><b>Código</b><span>${esc(codigoVisible)}</span></div><div class="d"><b>Inventario</b><span>${existencias} materiales · ${movimientos} movimientos históricos</span></div>${materialesHtml}${movimientos?`<p class="ubicaciones-ayuda">Al eliminarla, sus movimientos permanecerán en el historial con el nombre y código originales.</p>`:""}<div id="ubiQr-${i}"></div>${esAdmin?`<div class="ubicacion-admin-actions"><button type="button" class="bd ubicacion-delete" data-eliminar-ubicacion="${esc(loc.codigo)}" ${ocupada?`disabled title="${movimientosError?"No se pudo verificar el historial":"Traspasa el inventario antes de eliminar"}"`:""}>🗑️ Eliminar ubicación</button></div>`:""}</div></details>`
  }).join("");
  lista.querySelectorAll(".ubicacion-card").forEach(item=>{
    item.addEventListener("toggle",()=>{
      if(!item.open)return;
      lista.querySelectorAll(".ubicacion-card[open]").forEach(otra=>{if(otra!==item)otra.open=false});
      const i=Number(item.dataset.index);mostrarQrUbicacion(ubicaciones[i],i);
    });
  });
  lista.querySelectorAll(".ubicacion-card[open]").forEach(item=>{
    const i=Number(item.dataset.index);mostrarQrUbicacion(ubicaciones[i],i);
  });
  lista.querySelectorAll("[data-eliminar-ubicacion]").forEach(btn=>btn.onclick=()=>eliminarUbicacion(btn.dataset.eliminarUbicacion));
}
function codigoQrUbicacion(codigo){return`TRAZABILIDAD-UBICACION:${normalizarCodigoUbicacion(codigo)}`}
function generarEtiquetaQrUbicacion(loc){
  const qr=dibujarQR(codigoQrUbicacion(loc.codigo),600);
  if(!qr)return null;
  const almacen=almacenes.find(a=>a.codigo===loc.almacen_codigo),nombre=nombreUbicacion(loc);
  const canvas=document.createElement("canvas"),ctx=canvas.getContext("2d");
  if(!ctx)return null;
  canvas.width=800;canvas.height=1080;ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);
  ctx.imageSmoothingEnabled=false;ctx.drawImage(qr,100,35,600,600);
  const texto=(valor,y,tamano,peso="600")=>{
    const limite=720,lineas=[];let linea="";
    ctx.font=`${peso} ${tamano}px Arial,sans-serif`;
    String(valor||"").split(/\s+/).filter(Boolean).forEach(palabra=>{
      const prueba=linea?`${linea} ${palabra}`:palabra;
      if(ctx.measureText(prueba).width<=limite){linea=prueba;return}
      if(linea)lineas.push(linea);
      linea=palabra;
      while(ctx.measureText(linea).width>limite&&linea.length>1){let corte=linea.length-1;while(corte>1&&ctx.measureText(linea.slice(0,corte)).width>limite)corte--;lineas.push(linea.slice(0,corte));linea=linea.slice(corte)}
    });
    if(linea)lineas.push(linea);
    ctx.fillStyle="#111";ctx.textAlign="center";
    lineas.forEach((renglon,i)=>ctx.fillText(renglon,400,y+i*tamano*1.2,limite));
    return lineas.length*tamano*1.2;
  };
  texto(normalizarCodigoUbicacion(loc.codigo),690,48,"700");
  texto("UBICACIÓN",760,38,"700");
  let y=815;
  if(almacen?.nombre)y+=texto(almacen.nombre,y,30);
  if(nombre&&nombre!==normalizarCodigoUbicacion(loc.codigo))texto(nombre,y+8,32);
  return canvas;
}
function mostrarQrUbicacion(loc,i){
  const box=$(`ubiQr-${i}`);if(!box)return;
  if(box.querySelector("canvas"))return;
  const qr=dibujarQR(codigoQrUbicacion(loc.codigo),240);
  if(!qr){toast("No se pudo generar el QR de la ubicación",1);return}
  const almacen=almacenes.find(a=>a.codigo===loc.almacen_codigo);
  box.innerHTML=`<div class="acts"><button type="button" class="bs" data-ubi-descargar>⬇️ Descargar etiqueta QR</button><button type="button" class="bs" data-ubi-imprimir>🖨️ Imprimir QR</button></div>`;
  qr.className="ubicacion-qr";box.prepend(qr);
  box.querySelector("[data-ubi-descargar]").onclick=async()=>{
    const etiqueta=generarEtiquetaQrUbicacion(loc);
    if(!etiqueta){toast(`No se pudo generar la etiqueta QR de ${normalizarCodigoUbicacion(loc.codigo)}.`,1);return}
    await bajarBlob(`QR_${safeN(normalizarCodigoUbicacion(loc.codigo))}.png`,new Blob([durl(etiqueta.toDataURL("image/png"))],{type:"image/png"}));
  };
  box.querySelector("[data-ubi-imprimir]").onclick=()=>imprimirEtiquetaUbicacion(loc,qr.toDataURL("image/png"),almacen);
}
$("bUbiExportar").onclick=()=>{
  if(!esAdmin){toast("Solo el administrador puede exportar información.",1);return}
  if(!ubicaciones.length){toast("No hay ubicaciones para exportar.",1);return}
  const archivos=[];
  for(const loc of ubicaciones){
    const etiqueta=generarEtiquetaQrUbicacion(loc);
    if(!etiqueta){toast(`No se pudo generar la etiqueta QR de ${normalizarCodigoUbicacion(loc.codigo)}. No se exportó el archivo.`,1);return}
    archivos.push({name:`QR_${safeN(normalizarCodigoUbicacion(loc.codigo))}.png`,data:durl(etiqueta.toDataURL("image/png"))});
  }
  bajarBlob("QR_ubicaciones.zip",zipBytes(archivos));
  toast(`${archivos.length} etiquetas QR de ubicaciones exportadas ⬇️`);
};
function imprimirEtiquetaUbicacion(loc,src,almacen){
  const codigoVisible=normalizarCodigoUbicacion(loc.codigo),nombreVisible=nombreUbicacion(loc);
  const e=$("etiqueta");e.className="hoja";e.innerHTML=`<div class="pz"><img src="${src}" alt="QR ${esc(codigoVisible)}"><div class="tx"><b class="cd">${esc(codigoVisible)}</b><span class="tg">UBICACIÓN</span><small>${esc(almacen?.nombre||"Sin almacén")}</small>${nombreVisible!==codigoVisible?`<small>${esc(nombreVisible)}</small>`:""}</div></div>`;
  let st=$("etqCss");if(!st){st=document.createElement("style");st.id="etqCss";document.head.appendChild(st)}
  st.textContent=cssEtq();window.addEventListener("afterprint",limpiarEtqCss,{once:true});window.print();
}
function procesarEscaneoInventario(txt){
  const raw=String(txt||"").trim(),prefijo="TRAZABILIDAD-UBICACION:";
  if(raw.toUpperCase().startsWith(prefijo)){
    const codigo=raw.slice(prefijo.length).trim(),loc=ubicacionPorCodigo(codigo);
    if(!loc){toast(ubicacionesAliasError||`Ubicación no registrada: ${codigo}`,1);return}
    const almacen=almacenes.find(a=>a.codigo===loc.almacen_codigo);
    ubicacionDestino=loc;$("ubiDestino").textContent=`${almacen?.nombre||"Sin almacén"} → ${ubicacionConCodigo(loc)}`;renderAsignacion();toast(`Destino seleccionado: ${almacen?.nombre?almacen.nombre+" → ":""}${ubicacionConCodigo(loc)}`);$("ubiScan").focus();return
  }
  const cod=codDe(raw).toLowerCase(),m=mats.find(x=>x.codigo.toLowerCase()===cod);
  if(!m){toast(`Código no registrado: ${cod}`,1);return}
  if(!ubicacionDestino){toast("Primero escanea el QR de la ubicación destino.",1);return}
  if(asignacionItems.some(x=>x.codigo.toLowerCase()===m.codigo.toLowerCase())){toast(`${m.codigo} ya está en la lista de traspaso`,1);return}
  const e=existencia(m);if(e.rest<=0){toast(`${m.codigo} no tiene existencia disponible`,1);return}
  asignacionItems.push({codigo:m.codigo});renderAsignacion();
}
function renderAsignacion(){
  $("asigLista").innerHTML=asignacionItems.map((item,i)=>{const m=mats.find(x=>x.codigo===item.codigo);if(!m)return"";const origen=ubicaciones.find(loc=>loc.codigo===m.ubicacionCodigo),origenTxt=origen?ubicacionConCodigo(origen):m.ubicacion||"Sin ubicación";return`<div class="asig-item"><span><b>${esc(m.codigo)}</b> · ${esc(m.descripcion)} · ${esc(existTxt(m))}<br><small>Desde: ${esc(origenTxt)}</small></span><button type="button" class="bd" data-asig-quitar="${i}">Quitar</button></div>`}).join("");
  $("asigLista").querySelectorAll("[data-asig-quitar]").forEach(b=>b.onclick=()=>{asignacionItems.splice(+b.dataset.asigQuitar,1);renderAsignacion()});
  $("ubiConfirmar").disabled=!ubicacionDestino||!asignacionItems.length||!!ubicacionesError;
}
async function confirmarAsignacion(){
  if(!ubicacionDestino||!asignacionItems.length)return;
  const items=asignacionItems.map(x=>mats.find(m=>m.codigo===x.codigo)).filter(Boolean);
  if(items.length!==asignacionItems.length){toast("Uno o más materiales ya no están disponibles. Actualice e intente de nuevo.",1);await cargar();renderAsignacion();return}
  const mismo=items.filter(m=>m.ubicacionCodigo===ubicacionDestino.codigo||(!m.ubicacionCodigo&&m.ubicacion===ubicacionDestino.nombre&&(!m.ubicacionAlmacenCodigo||m.ubicacionAlmacenCodigo===ubicacionDestino.almacen_codigo)));
  if(mismo.length){toast(`Ya están en ${ubicacionConCodigo(ubicacionDestino)}: ${mismo.map(m=>m.codigo).join(", ")}`,1);return}
  const folio=`UBI-${Date.now()}`;
  const payload=items.map(m=>{const e=existencia(m);return{codigo:m.codigo,existencia:{modo:e.modo,cantidad:e.modo==="pz"?e.rest:e.qty,longitud:e.modo==="long"?e.rest:null,unidad:e.u,peso:e.peso?.peso??null,pesoUnidad:e.peso?.u??null}}});
  $("ubiConfirmar").disabled=true;
  try{
    const{data,error}=await sb.rpc("asignar_inventario_ubicacion",{p_ubicacion_codigo:ubicacionDestino.codigo,p_folio:folio,p_items:payload});
    if(error)throw error;
    if(Number(data)!==items.length)throw new Error(`Respuesta incompleta del servidor (${data??"sin cantidad"})`);
    asignacionItems=[];renderAsignacion();
    try{await Promise.all([cargar(),cargarUbicaciones()])}
    catch(err){errDB(err,"Traspaso registrado, pero no se pudo actualizar la vista")}
    toast(`Traspaso registrado a ${ubicacionConCodigo(ubicacionDestino)} · ${folio}`);
  }catch(err){errDB(err,"No se pudo registrar el traspaso");$("ubiConfirmar").disabled=false}
}
function renderHistUbicaciones(){
  $("ubiHistEstado").innerHTML=[movimientosError,ajustesInventarioError].filter(Boolean).map(error=>`<div class="res bad">${esc(error)}</div>`).join("");
  const filas=movimientosUbicacion.slice().sort((a,b)=>new Date(b.created_at)-new Date(a.created_at));
  $("ubiHistBody").innerHTML=filas.map(x=>{const anterior=almacenes.find(a=>a.codigo===x.ubicacion_anterior_almacen_codigo),nueva=almacenes.find(a=>a.codigo===x.ubicacion_nueva_almacen_codigo),desde=`${esc(x.ubicacion_anterior||"Sin ubicación")}${x.ubicacion_anterior_codigo?` <small>(${esc(x.ubicacion_anterior_codigo)})</small>`:""}`,hacia=`${esc(x.ubicacion_nueva)}${x.ubicacion_nueva_codigo?` <small>(${esc(x.ubicacion_nueva_codigo)})</small>`:""}`;return`<tr><td>${esc(new Date(x.created_at).toLocaleString("es-MX"))}</td><td>${esc(x.folio)}</td><td>${esc(x.codigo_material)}</td><td>${esc(anterior?.nombre||"Sin almacén")}</td><td>${desde}</td><td>${esc(nueva?.nombre||"Sin almacén")}</td><td>${hacia}</td><td>${esc(resumenExistencia(x.existencia))}</td></tr>`}).join("");
  $("ubiHistVacio").hidden=filas.length>0;
  const conteos=mats.flatMap(m=>(m.ajustesInventario||[]).map(ajuste=>({...ajuste,codigo_material:m.codigo}))).sort((a,b)=>new Date(b.created_at)-new Date(a.created_at));
  $("conteoHistBody").innerHTML=conteos.map(x=>`<tr><td>${esc(new Date(x.created_at).toLocaleString("es-MX"))}</td><td><b>${esc(x.codigo_material)}</b></td><td>${esc(x.ubicacion_nombre||x.ubicacion_codigo||"")}</td><td>${esc(fmt(Number(x.existencia_anterior)))} ${esc(x.unidad)}</td><td>${esc(fmt(Number(x.existencia_contada)))} ${esc(x.unidad)}</td><td>${esc(fmt(Number(x.diferencia)))} ${esc(x.unidad)}</td><td class="l">${esc(x.motivo||"—")}</td><td>${esc(x.usuario||"—")}</td></tr>`).join("");
  $("conteoHistVacio").hidden=conteos.length>0;
}
function resumenExistencia(e){
  if(!e)return"";
  const cantidad=e.modo==="long"?`${fmt(e.longitud)} ${e.unidad||"mm"}`:`${fmt(e.cantidad)} ${e.unidad||"PZ"}`;
  return`${cantidad}${e.peso!=null?` · ${fmt(e.peso)} ${e.pesoUnidad}`:""}`;
}
function siguienteCodigoUbicacion(){
  const max=ubicaciones.reduce((n,l)=>Math.max(n,Number(/^(?:UBI-)?(\d+)$/i.exec(l.codigo)?.[1])||0),0);
  $("ubiCodigo").value=`UBI-${String(max+1).padStart(3,"0")}`;
}
async function resolverUbicacionRegistro(nombre,tipoRegistro,ubicacionAnteriorCodigo){
  const almacenCodigo={MP:"61",CONSUMIBLES:"62"}[tipoRegistro],valor=String(nombre||"").trim(),clave=valor.toLocaleLowerCase(),codigoIngresado=normalizarCodigoUbicacion(valor).toLocaleLowerCase(),codigoManual=/^(?:UBI-)?\d+$/i.test(valor),codigoSolicitado=normalizarCodigoUbicacion(valor);
  if(!clave)throw new Error("La ubicación es obligatoria.");
  if(almacenesError||ubicacionesError)throw new Error("No se pudieron consultar los almacenes y ubicaciones. Recarga antes de guardar.");
  const coincideNombre=loc=>[loc.nombre,loc.codigo].some(texto=>String(texto||"").trim().toLocaleLowerCase()===clave||normalizarCodigoUbicacion(texto).toLocaleLowerCase()===codigoIngresado);
  const anterior=ubicaciones.find(loc=>loc.codigo===ubicacionAnteriorCodigo&&coincideNombre(loc));
  let almacen=almacenes.find(a=>a.codigo===almacenCodigo);
  if(!almacenCodigo){
    if(anterior)almacen=almacenes.find(a=>a.codigo===anterior.almacen_codigo);
    const coincidencias=ubicaciones.filter(coincideNombre);
    const existente=anterior|| (coincidencias.length===1?coincidencias[0]:null);
    if(!existente)throw new Error(coincidencias.length>1?"La ubicación existe en más de un almacén. Selecciona Materia prima o Consumibles para indicar el almacén correcto.":"Para crear una ubicación desde el registro, selecciona Materia prima (almacén 61) o Consumibles (almacén 62).");
    almacen=almacenes.find(a=>a.codigo===existente.almacen_codigo);
    if(!almacen)throw new Error("No se encontró el almacén de la ubicación seleccionada.");
    return{ubicacion:existente,almacen,creada:false};
  }
  if(!almacen)throw new Error(`No se encontró el almacén ${almacenCodigo} para ${nombreTipoRegistro(tipoRegistro)}.`);
  const existentePorCodigo=ubicaciones.find(loc=>loc.almacen_codigo===almacenCodigo&&normalizarCodigoUbicacion(loc.codigo).toLocaleLowerCase()===codigoIngresado);
  if(existentePorCodigo)return{ubicacion:existentePorCodigo,almacen,creada:false};
  if(ubicacionesAlias.some(alias=>alias.codigo_alias.toLocaleLowerCase()===clave||normalizarCodigoUbicacion(alias.codigo_alias).toLocaleLowerCase()===codigoIngresado))
    throw new Error(`El código ${valor} está reservado como QR anterior de otra ubicación.`);
  const existente=ubicaciones.find(loc=>loc.almacen_codigo===almacenCodigo&&coincideNombre(loc));
  if(codigoManual&&existente){
    if(!esAdmin)throw new Error(`La ubicación ${valor} ya existe con el código ${existente.codigo}. Pide al administrador que corrija esa ubicación.`);
    const{data,error}=await sb.rpc("corregir_codigo_ubicacion",{
      p_codigo_actual:existente.codigo,p_codigo_nuevo:codigoSolicitado,p_almacen_codigo:almacenCodigo
    });
    if(error){
      if(/PGRST202|function.*does not exist|schema cache/i.test(String(error.message||"")))
        throw new Error("Falta aplicar la versión actualizada de supabase_almacenes.sql para corregir esta ubicación.");
      throw error;
    }
    if(data?.codigo_nuevo!==codigoSolicitado)throw new Error("El servidor no confirmó la corrección del código de ubicación.");
    const corregida={...existente,codigo:codigoSolicitado,nombre:codigoSolicitado};
    ubicaciones=ubicaciones.map(loc=>loc.codigo===existente.codigo?corregida:loc);
    ubicacionesAlias=ubicacionesAlias.map(alias=>alias.ubicacion_codigo===existente.codigo?{...alias,ubicacion_codigo:codigoSolicitado}:alias);
    if(!ubicacionesAlias.some(alias=>alias.codigo_alias===existente.codigo))ubicacionesAlias.push({codigo_alias:existente.codigo,ubicacion_codigo:codigoSolicitado});
    mats.forEach(material=>{
      if(material.ubicacionCodigo!==existente.codigo)return;
      material.ubicacionCodigo=codigoSolicitado;material.ubicacion=codigoSolicitado;
      material.ubicacionAlmacenCodigo=almacenCodigo;material.ubicacionAlmacenNombre=almacen.nombre;
    });
    $("ubiAlmacen").value=almacenCodigo;renderAlmacenes();renderUbicaciones();
    const indice=ubicaciones.indexOf(corregida),tarjeta=[...$("ubiLista").querySelectorAll(".ubicacion-card")].find(item=>item.dataset.codigo===codigoSolicitado);
    if(tarjeta){tarjeta.open=true;mostrarQrUbicacion(corregida,indice)}
    return{ubicacion:corregida,almacen,creada:false,corregida:true};
  }
  if(existente)return{ubicacion:existente,almacen,creada:false};
  let codigo=codigoManual?codigoSolicitado:valor;
  if(!codigoManual){
    const max=ubicaciones.reduce((n,loc)=>Math.max(n,Number(/^(?:UBI-)?(\d+)$/i.exec(loc.codigo)?.[1])||0),0);
    let numero=max+1;codigo=`UBI-${String(numero).padStart(3,"0")}`;
    while(ubicaciones.some(loc=>loc.codigo.toLocaleLowerCase()===codigo.toLocaleLowerCase())||ubicacionesAlias.some(alias=>alias.codigo_alias.toLocaleLowerCase()===codigo.toLocaleLowerCase())){numero++;codigo=`UBI-${String(numero).padStart(3,"0")}`}
  }
  const{data,error}=await sb.from("ubicaciones").insert({codigo,nombre:valor,almacen_codigo:almacenCodigo}).select("codigo,nombre,almacen_codigo,created_at").single();
  if(error){
    if(error.code==="23505"){
      const actualizadas=await leerTodo("ubicaciones","codigo,nombre,almacen_codigo,created_at",["nombre"]);
      ubicaciones=actualizadas;
      const concurrente=ubicaciones.find(loc=>loc.almacen_codigo===almacenCodigo&&coincideNombre(loc));
      if(concurrente){$("ubiAlmacen").value=almacenCodigo;renderAlmacenes();renderUbicaciones();return{ubicacion:concurrente,almacen,creada:false}}
    }
    throw error;
  }
  if(!data)throw new Error("El servidor no devolvió la ubicación registrada.");
  ubicaciones.push(data);$("ubiAlmacen").value=almacenCodigo;renderAlmacenes();renderUbicaciones();
  const indice=ubicaciones.indexOf(data),tarjeta=[...$("ubiLista").querySelectorAll(".ubicacion-card")].find(item=>item.dataset.codigo===data.codigo);
  if(tarjeta){tarjeta.open=true;mostrarQrUbicacion(data,indice)}
  return{ubicacion:data,almacen,creada:true};
}
function siguienteCodigoAlmacen(){
  const max=almacenes.reduce((n,a)=>Math.max(n,Number(/^ALM-(\d+)$/i.exec(a.codigo)?.[1])||0),0);
  $("almacenCodigo").value=`ALM-${String(max+1).padStart(3,"0")}`;
}
$("almacenAuto").onclick=siguienteCodigoAlmacen;
$("almacenForm").addEventListener("submit",async e=>{
  e.preventDefault();
  const codigo=$("almacenCodigo").value.trim(),nombre=$("almacenNombre").value.trim(),clasificacion=$("almacenTipo").value||null;
  if(!codigo||!nombre)return;
  $("almacenGuardar").disabled=true;
  try{
    const{error}=await sb.from("almacenes").insert({codigo,nombre,clasificacion});if(error)throw error;
    $("almacenNombre").value="";$("almacenTipo").value="";siguienteCodigoAlmacen();
    await cargarUbicaciones();$("ubiAlmacen").value=codigo;renderUbicaciones();
    toast(`Almacén guardado: ${nombre}`);
  }catch(err){errDB(err,"No se pudo guardar el almacén")}
  finally{$("almacenGuardar").disabled=false}
});
$("ubiAuto").onclick=siguienteCodigoUbicacion;
$("ubiAlmacen").onchange=renderUbicaciones;
function leerTextoArchivo(f){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result||""));r.onerror=rej;r.readAsText(f,"utf-8")})}
function normalizarClave(c){return String(c||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]/g,"")}
function buscarValorFila(obj,claves){if(!obj)return"";const mapa=new Map(Object.entries(obj).map(([k,v])=>[normalizarClave(k),v]));for(const c of claves){const k=normalizarClave(c),v=mapa.get(k);if(v!==undefined&&v!==null&&String(v).trim()!=="")return v}return""}
function obtenerAlmacenCodigo(valor){if(!valor)return"";const v=String(valor).trim();if(!v)return"";const a=almacenes.find(x=>x.codigo.toLowerCase()===v.toLowerCase()||x.nombre.toLowerCase()===v.toLowerCase());return a?a.codigo:v}
async function importarUbicacionesExcel(file){
  if(!esAdmin)throw new Error("Solo el administrador puede importar ubicaciones.");
  await cargarXLSX();
  if(almacenesError||ubicacionesError)throw new Error("Corrija los errores de almacenes/ubicaciones antes de importar.");
  const workbook=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true}),nombreHoja=workbook.SheetNames[0],sheet=workbook.Sheets[nombreHoja];
  if(!sheet)throw new Error("El archivo no contiene hojas para importar.");
  const rows=XLSX.utils.sheet_to_json(sheet,{defval:"",raw:true});
  if(!rows.length)throw new Error("No se encontraron filas para importar.");
  const existentes=new Set(ubicaciones.map(u=>normalizarCodigoUbicacion(u.codigo).toLowerCase()));
  const porAlmacen=new Map(almacenes.map(a=>[a.codigo.toLowerCase(),a.codigo]));
  const aInsertar=[],omitidos=[],sinAlmacen=[],duplicados=[];
  rows.forEach(row=>{
    const valor=String(buscarValorFila(row,["codigo","ubicacion","codigo_ubicacion","codigoubicacion","ubi","clave"])).trim();
    if(!valor)return;
    const codigo=normalizarCodigoUbicacion(valor);
    if(existentes.has(codigo.toLowerCase())){duplicados.push(codigo);return}
    const almacenValor=buscarValorFila(row,["almacen_codigo","almacencodigo","almacen","clave_almacen","almacen_id"]);
    const almacenCodigo=obtenerAlmacenCodigo(almacenValor);
    if(!almacenCodigo){sinAlmacen.push(codigo);return}
    if(!porAlmacen.has(almacenCodigo.toLowerCase())){sinAlmacen.push(codigo);return}
    aInsertar.push({codigo,nombre:codigo,almacen_codigo:almacenCodigo});
    existentes.add(codigo.toLowerCase());
  });
  if(!aInsertar.length){throw new Error(`No se importaron ubicaciones. Duplicados: ${duplicados.length} · Sin almacén válido: ${sinAlmacen.length}`)}
  const{error}=await sb.from("ubicaciones").insert(aInsertar);
  if(error)throw error;
  await cargarUbicaciones();
  toast(`Ubicaciones importadas: ${aInsertar.length} · Duplicados: ${duplicados.length} · Sin almacén válido: ${sinAlmacen.length}`);
}
async function importarUbicacionesJson(file){
  if(!esAdmin)throw new Error("Solo el administrador puede importar ubicaciones.");
  if(almacenesError||ubicacionesError)throw new Error("Corrija los errores de almacenes/ubicaciones antes de importar.");
  const texto=await leerTextoArchivo(file);
  let data;try{data=JSON.parse(texto)}catch{throw new Error("Archivo JSON no válido")}
  const filas=Array.isArray(data)?data:Array.isArray(data?.ubicaciones)?data.ubicaciones:Array.isArray(data?.items)?data.items:[];
  if(!filas.length)throw new Error("No se encontraron ubicaciones en el JSON.");
  const existentes=new Set(ubicaciones.map(u=>normalizarCodigoUbicacion(u.codigo).toLowerCase()));
  const porAlmacen=new Map(almacenes.map(a=>[a.codigo.toLowerCase(),a.codigo]));
  const aInsertar=[],duplicados=[],sinAlmacen=[];
  filas.forEach(row=>{
    const valor=String(buscarValorFila(row,["codigo","ubicacion","codigo_ubicacion","ubi"])).trim();
    if(!valor)return;
    const codigo=normalizarCodigoUbicacion(valor);
    if(existentes.has(codigo.toLowerCase())){duplicados.push(codigo);return}
    const almacenCodigo=obtenerAlmacenCodigo(buscarValorFila(row,["almacen_codigo","almacencodigo","almacen","almacen_id"]));
    if(!almacenCodigo||!porAlmacen.has(almacenCodigo.toLowerCase())){sinAlmacen.push(codigo);return}
    aInsertar.push({codigo,nombre:codigo,almacen_codigo:almacenCodigo});
    existentes.add(codigo.toLowerCase());
  });
  if(!aInsertar.length){throw new Error(`No se importaron ubicaciones. Duplicados: ${duplicados.length} · Sin almacén válido: ${sinAlmacen.length}`)}
  const{error}=await sb.from("ubicaciones").insert(aInsertar);
  if(error)throw error;
  await cargarUbicaciones();
  toast(`Ubicaciones importadas: ${aInsertar.length} · Duplicados: ${duplicados.length} · Sin almacén válido: ${sinAlmacen.length}`);
}
$("bUbiExcel").onclick=()=>{if(!esAdmin){toast("Solo el administrador puede importar ubicaciones.",1);return}$("fileUbiExcel").click()};
$("bUbiJson").onclick=()=>{if(!esAdmin){toast("Solo el administrador puede importar ubicaciones.",1);return}$("fileUbiJson").click()};
$("fileUbiExcel").onchange=async e=>{const file=e.target.files[0];e.target.value="";if(!file)return;try{await importarUbicacionesExcel(file)}catch(err){toast(err.message||"No se pudo importar ubicaciones",1)}};
$("fileUbiJson").onchange=async e=>{const file=e.target.files[0];e.target.value="";if(!file)return;try{await importarUbicacionesJson(file)}catch(err){toast(err.message||"No se pudo importar ubicaciones",1)}};
$("ubiForm").addEventListener("submit",async e=>{
  e.preventDefault();const codigo=normalizarCodigoUbicacion($("ubiCodigo").value);
  const almacenCodigo=$("ubiAlmacen").value,almacen=almacenes.find(a=>a.codigo===almacenCodigo);
  if(!codigo||!almacen)return;
  if(ubicaciones.some(loc=>loc.almacen_codigo===almacenCodigo&&normalizarCodigoUbicacion(loc.codigo).toLocaleLowerCase()===codigo.toLocaleLowerCase())){
    toast(`La ubicación ${codigo} ya está registrada en ${almacen.nombre}.`,1);return;
  }
  $("ubiGuardar").disabled=true;
  try{const{error}=await sb.from("ubicaciones").insert({codigo,nombre:codigo,almacen_codigo:almacenCodigo});if(error)throw error;
    $("ubiCodigo").value="";siguienteCodigoUbicacion();await cargarUbicaciones();$("ubiAlmacen").value=almacenCodigo;renderUbicaciones();toast(`Ubicación guardada en ${almacen.nombre}: ${codigo}`);
  }catch(err){errDB(err,"No se pudo guardar la ubicación")}
  finally{$("ubiGuardar").disabled=false}
});
function manejarLectorUbicacion(){
  const field=$("ubiScan"),valor=field.value.trim();if(!valor)return;
  procesarEscaneoInventario(valor);field.value="";
}
$("ubiScan").addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();manejarLectorUbicacion()}});
$("ubiConfirmar").onclick=confirmarAsignacion;
$("ubiLimpiar").onclick=()=>{asignacionItems=[];ubicacionDestino=null;$("ubiDestino").textContent="Sin seleccionar";renderAsignacion()};
function syncBulkActions(){
  const visible=[...$("tbody").querySelectorAll("[data-select]")],checked=visible.filter(c=>c.checked).length,all=$("selectAll");
  all.hidden=!esAdmin;$("seleccionInfo").hidden=!esAdmin;$("bDelSel").hidden=$("bDelAll").hidden=!esAdmin;
  all.checked=visible.length>0&&checked===visible.length;all.indeterminate=checked>0&&checked<visible.length;
  $("seleccionInfo").textContent=`${seleccionados.size} seleccionados`;
  $("bDelSel").disabled=seleccionados.size===0;$("bDelAll").disabled=mats.length===0;
}
$("buscar").oninput=deb(render);$("ftipo").onchange=render;
document.querySelectorAll("th[data-k]").forEach(th=>th.onclick=()=>{orden={k:th.dataset.k,asc:orden.k===th.dataset.k?!orden.asc:true};render()});
$("selectAll").addEventListener("change",e=>{
  $("tbody").querySelectorAll("[data-select]").forEach(c=>{c.checked=e.target.checked;if(c.checked)seleccionados.add(c.dataset.select);else seleccionados.delete(c.dataset.select)});
  syncBulkActions();
});
$("tbody").addEventListener("change",e=>{
  const c=e.target.closest("[data-select]");if(!c)return;
  if(c.checked)seleccionados.add(c.dataset.select);else seleccionados.delete(c.dataset.select);
  syncBulkActions();
});
function alternarQrMaterial(m){
  const panel=$("vReg").querySelector('[data-module-screen="qr"]');
  if(actual?.codigo===m.codigo&&!panel.hidden){
    panel.hidden=true;actual=null;
    $("vReg").classList.remove("qr-visible");
    $("qrBox").innerHTML="Guarde o seleccione un material para generar su QR";
    $("ficha").innerHTML="";
    $("bPrint").disabled=$("bPng").disabled=true;
    $("piezasBox").hidden=true;
    $("pzGrid").innerHTML="";
    return;
  }
  mostrar(m);
}
$("tbody").addEventListener("click",async e=>{
  const b=e.target.closest("button");
  if(b){
    const i=+b.dataset.i,m=mats[i];
    if(b.dataset.a==="qr"){alternarQrMaterial(m);return}
    if(b.dataset.a==="ed"){if(esAdmin)editar(i);else toast("Solo el administrador puede editar materiales.",1);return}
    if(b.dataset.a==="del"){
      if(!esAdmin){toast("Solo el administrador puede eliminar materiales.",1);return}
      if(await confirmar({titulo:"Eliminar material",mensaje:`¿Eliminar ${m.codigo}? Se borrarán también sus ${(m.despachos||[]).length} surtido(s).`,boton:"Eliminar",peligro:true})){try{const{error}=await sb.from("materiales").delete().eq("codigo",m.codigo);if(error)throw error;await cargar();toast("Material eliminado")}catch(err){errDB(err,"No se pudo eliminar")}}
    }
    return;
  }
  if(e.target.closest("input,select,textarea,a"))return;
  const fila=e.target.closest("tr"),botonQr=fila?.querySelector('button[data-a="qr"]');
  if(!botonQr)return;
  const m=mats[+botonQr.dataset.i];
  if(m)alternarQrMaterial(m);
});
async function eliminarEnLotes(codigos){
 if(!esAdmin)throw new Error("Solo el administrador puede eliminar materiales.");
 for(let i=0;i<codigos.length;i+=100){
   const{error}=await sb.from("materiales").delete().in("codigo",codigos.slice(i,i+100));
   if(error)throw error;
 }
}
$("bDelSel").onclick=async()=>{
 if(!esAdmin){toast("Solo el administrador puede eliminar materiales.",1);return}
 const codigos=[...seleccionados].filter(c=>mats.some(m=>m.codigo===c));
 if(!codigos.length){seleccionados.clear();render();return}
 if(!await confirmar({titulo:"Eliminar materiales",mensaje:`Se eliminarán ${codigos.length} material(es) seleccionado(s) y sus surtidos.`,escribir:codigos.length>5?"ELIMINAR":"",boton:"Eliminar",peligro:true}))return;
 $("bDelSel").disabled=$("bDelAll").disabled=true;
 try{await eliminarEnLotes(codigos);codigos.forEach(c=>seleccionados.delete(c));toast(`${codigos.length} material(es) eliminado(s)`)}
 catch(err){errDB(err,"No se pudieron eliminar todos los materiales")}
 try{await cargar()}catch(err){errDB(err,"No se pudo actualizar la lista")}
};
$("bDelAll").onclick=async()=>{
 if(!esAdmin){toast("Solo el administrador puede restablecer el catálogo.",1);return}
 const codigos=mats.map(m=>m.codigo);
 if(!codigos.length)return;
 if(!await confirmar({titulo:"Eliminar TODO el catálogo",mensaje:`Se eliminarán ${codigos.length} materiales y todos sus surtidos. Esta acción no se puede deshacer.`,escribir:"ELIMINAR TODO",boton:"Eliminar todo",peligro:true}))return;
 $("bDelSel").disabled=$("bDelAll").disabled=true;
 try{await eliminarEnLotes(codigos);seleccionados.clear();toast(`${codigos.length} material(es) eliminado(s)`)}
 catch(err){errDB(err,"No se pudieron eliminar todos los materiales")}
 try{await cargar()}catch(err){errDB(err,"No se pudo actualizar la lista")}
};
function editar(i){
  if(!esAdmin){toast("Solo el administrador puede editar materiales.",1);return}
  activarSubpantalla("vReg","captura");
  const m=mats[i];resetForm();$("editIndex").value=i;
  const values={codigo:m.codigo,descripcion:m.descripcion,um:m.um,cantidad:m.cantidad,ubicacion:m.ubicacion,fechaRecibo:m.fechaRecibo,tipoRegistro:m.tipoRegistro|| (m.esMP?"MP":"GENERAL"),seccionMaterial:m.noAplica?.seccionMaterial?"":m.seccionMaterial,altoMaterial:m.alto,anchoMaterial:m.ancho,diametroMaterial:m.diametro,
    tipoMaterial:m.tipoMaterial??(m.esMP?"Materia prima":m.tipoRegistro==="CONSUMIBLES"?"Consumibles":"Material general"),especificacionMedida:m.especificacionMedida,longitudIndividual:m.longitudIndividual,manufactura:m.manufactura,
    numeroHeat:m.numeroHeat,numeroLoteMP:m.numeroLoteMP??m.lote,numeroCertificado:m.numeroCertificado,totalMM:m.totalMM,peso:m.peso,ordenMaterial:m.ordenMaterial,poNumber:m.poNumber};
  Object.entries(values).forEach(([id,value])=>$(id).value=value??"");
  actualizarCamposSeccion();
  $("um").value="PZ";
  $("pesoU").value=m.pesoU||m.pesoUnidad||(m.pesoUnidadPendiente?"":(["LB","KG"].includes(m.um)?m.um:"LB"));
  CAMPOS_EXTRA.forEach(([key,id])=>{
    const noAplica=!!m.noAplica?.[key],checkbox=document.querySelector(`input[data-na="${id}"]`),input=$(id);
    checkbox.checked=noAplica;input.value=values[key]??"";
    const wrap=input.closest(".na-field").querySelector(".na-value-wrap")||input;wrap.hidden=noAplica;
  });
  const seccionNoAplica=!!m.noAplica?.seccionMaterial,seccionCheckbox=document.querySelector('input[data-na="seccionMaterial"]');
  seccionCheckbox.checked=seccionNoAplica;$("seccionMaterial").hidden=seccionNoAplica;
  actualizarCamposSeccion();
  $("tituloForm").textContent=`Editando: ${m.codigo}`;$("btnGuardar").textContent="💾 Guardar cambios y regenerar QR";$("btnCancelar").hidden=false;
  establecerFormularioAbierto(true);window.scrollTo({top:0,behavior:"smooth"});
  $("detTraza").open=true;
}

// --- Exportar / Importar ---
async function bajar(nombre,contenido,tipo){
  if(matchMedia("(pointer:coarse)").matches){try{const f=new File([contenido],nombre,{type:tipo.split(";")[0]});if(navigator.canShare&&navigator.canShare({files:[f]})){await navigator.share({files:[f],title:nombre});return}}catch(e){if(e.name==="AbortError")return}}
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([contenido],{type:tipo}));a.download=nombre;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
$("bJson").onclick=()=>{if(!esAdmin){toast("Solo el administrador puede exportar información.",1);return}bajar("trazabilidad_materiales.json",JSON.stringify(mats,null,2),"application/json");toast("JSON exportado ⬇️")};
$("bCsv").onclick=()=>{
  if(!esAdmin){toast("Solo el administrador puede exportar información.",1);return}
  const cols=["UBICACION","CODIGO","CLASIFICACION","TIPO MATERIAL","DESCRIPCION","ESPECIFICACION MEDIDA","SECCION","ALTO (MM)","ANCHO (MM)","DIAMETRO (MM)","LONGITUD IND (MM)","MANUFACTURA","N°HEAT","N°LOTE MP","N°CERT.","TOTAL MM","FECHA DE RECIBO","CTD","PESO","UM","ORDEN","PO NUMBER"];
  const c=v=>{let t=String(v??"");if(/^[=+\-@\t\r]/.test(t))t="'"+t;return `"${t.replace(/"/g,'""')}"`};
  const value=(m,key,v)=>m.noAplica?.[key]?"No aplica":v;
  const rows=mats.map(m=>[m.ubicacion,m.codigo,nombreTipoRegistro(m.tipoRegistro).toUpperCase(),value(m,"tipoMaterial",m.tipoMaterial),m.descripcion,
    value(m,"especificacionMedida",m.especificacionMedida),value(m,"seccionMaterial",m.seccionMaterial),m.alto,m.ancho,m.diametro,value(m,"longitudIndividual",m.longitudIndividual),value(m,"manufactura",m.manufactura),value(m,"numeroHeat",m.numeroHeat),
    value(m,"numeroLoteMP",m.numeroLoteMP||m.lote),value(m,"numeroCertificado",m.numeroCertificado),value(m,"totalMM",m.totalMM),
    m.fechaRecibo,m.cantidad,value(m,"peso",m.peso),m.noAplica?.peso?"No aplica":m.pesoUnidadPendiente?"Sin especificar":m.pesoU||m.pesoUnidad||"LB",
    value(m,"ordenMaterial",m.ordenMaterial),value(m,"poNumber",m.poNumber)]);
  bajar("trazabilidad_materiales.csv","\ufeff"+[cols.join(","),...rows.map(row=>row.map(c).join(","))].join("\r\n"),"text/csv;charset=utf-8");toast("CSV exportado ⬇️");
};
$("bImp").onclick=()=>{if(!esAdmin){toast("Solo el administrador puede restaurar información.",1);return}$("fileImp").click()};
$("bExcel").onclick=()=>{if(!esAdmin){toast("Solo el administrador puede importar información.",1);return}$("fileExcel").click()};
const EXCEL_HEADERS={
  ubicacion:["ubicacion","location"],codigo:["codigo","code"],tipoRegistro:["clasificacion","tiporegistro","categoriaregistro"],tipoMaterial:["tipomaterial","tipo"],descripcion:["descripcion","description"],
  seccionMaterial:["seccion","seccionmaterial","forma","shape"],alto:["alto","altomm","height"],ancho:["ancho","anchomm","width"],diametro:["diametro","diametromm","diameter"],
  especificacionMedida:["especificacionmedida","especificacion","medida"],longitudIndividual:["longitudind","longitudindmm","longituddelmaterial","longitudinddelmaterial","longitudindividual","longitudindividualmm","longitudindividualdelmaterial","longitudpza","longitudpieza","longitudporpieza","longind","longindmm"],manufactura:["manufactura","fabricacion"],
  numeroHeat:["nheat","numeroheat","heat"],numeroLoteMP:["nlotemp","numerolotemp","lotemp"],
  numeroCertificado:["ncert","ncertificado","numerocertificado","certificado"],totalMM:["totalmm"],
  fechaRecibo:["fechaderecibo","fecharecibo","fecha"],cantidad:["ctd","cantidad","quantity"],
  peso:["peso","weight"],um:["um","unidadmedida","unidad"],
  ordenMaterial:["orden","ordenmaterial","order"],poNumber:["ponumber","po"]
};
function normalizarEncabezado(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]/g,"")}
function buscarEncabezados(rows){
  for(let i=0;i<Math.min(rows.length,20);i++){
    const headers=rows[i].map(normalizarEncabezado),mapped={};
    Object.entries(EXCEL_HEADERS).forEach(([key,aliases])=>{const index=headers.findIndex(h=>aliases.includes(h));if(index>=0)mapped[key]=index});
    if(["codigo","descripcion","ubicacion","cantidad","fechaRecibo"].every(key=>mapped[key]!=null))return{index:i,mapped};
  }
  throw new Error("No se encontraron los encabezados requeridos: UBICACION, CODIGO, DESCRIPCION, FECHA DE RECIBO y CTD.");
}
function valorCelda(row,mapped,key){return mapped[key]==null?"":row[mapped[key]]??""}
function numeroExcel(value){
  if(typeof value==="number")return Number.isFinite(value)?value:null;
  let s=String(value??"").trim().replace(/\s/g,"");if(!s)return null;
  if(s.includes(",")&&s.includes("."))s=s.lastIndexOf(",")>s.lastIndexOf(".")?s.replace(/\./g,"").replace(",","."):s.replace(/,/g,"");
  else if(s.includes(","))s=s.replace(",",".");
  const n=Number(s);return Number.isFinite(n)?n:null;
}
function fechaExcel(value){
  if(value instanceof Date&&!isNaN(value))return value.toISOString().slice(0,10);
  if(typeof value==="number"){const d=XLSX.SSF.parse_date_code(value);return d?`${String(d.y).padStart(4,"0")}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`:""}
  const s=String(value??"").trim();let m;
  if(m=/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(s))return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  if(m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))return`${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  return"";
}
function esNoAplica(value){return["na","n/a","n.a.","no aplica","noaplica","no aplica."].includes(String(value??"").trim().toLowerCase())}
function normalizarSeccion(value){
  const s=String(value??"").trim().toUpperCase();
  if(["SQ","SQUARE","CUADRADA","CUADRADO"].includes(s))return"SQ";
  if(["FLAT","PLANA","PLANO"].includes(s))return"FLAT";
  if(["RND","ROUND","REDONDA","REDONDO","RD"].includes(s))return"RND";
  return"";
}
function inferirSeccion(texto){
  const s=String(texto??"").toUpperCase();
  if(/\b(?:SQ|SQUARE|CUADRADA|CUADRADO)\b/.test(s))return"SQ";
  if(/\b(?:FLAT|PLANA|PLANO)\b/.test(s))return"FLAT";
  if(/\b(?:RND|ROUND|REDONDA|REDONDO)\b/.test(s))return"RND";
  return"";
}
function materialDesdeExcel(row,mapped,esMateriaPrima=false){
  const get=key=>valorCelda(row,mapped,key),text=key=>String(get(key)??"").trim(),raw={};
  Object.keys(EXCEL_HEADERS).forEach(key=>raw[key]=get(key));
  const noAplica=Object.fromEntries(CAMPOS_EXTRA.map(([key])=>[key,esNoAplica(raw[key])]));
  noAplica.seccionMaterial=esNoAplica(get("seccionMaterial"));
  if(esNoAplica(get("um")))noAplica.peso=true;
  const stringValue=key=>noAplica[key]?"":String(raw[key]??"").trim();
  const codigo=text("codigo"),descripcion=text("descripcion"),ubicacion=text("ubicacion");
  const cantidad=numeroExcel(get("cantidad")),peso=numeroExcel(get("peso")),umOriginal=text("um").toUpperCase();
  const pesoUnitMap={LB:"LB",LBS:"LB",LIBRA:"LB",LIBRAS:"LB",KG:"KG",KGS:"KG",KILOGRAM:"KG",KILOGRAMO:"KG",KILOGRAMOS:"KG"};
  const umCantidad=["PZ","PZA","PIEZA","PIEZAS"].includes(umOriginal);
  const tipoArchivo=text("tipoRegistro").toUpperCase(),esMP=/^(?:MP|MATERIA PRIMA)$/.test(tipoArchivo)||esMateriaPrima||/materia\s*prima|^mp$/i.test(text("tipoMaterial")),esConsumible=/^CONSUMIBLES?$/.test(tipoArchivo);
  const seccionMaterial=noAplica.seccionMaterial?"":normalizarSeccion(get("seccionMaterial"))||inferirSeccion(`${text("especificacionMedida")} ${text("descripcion")}`);
  const fechaRecibo=fechaExcel(get("fechaRecibo"));
  if(!codigo||!descripcion||cantidad==null||cantidad<0||!fechaRecibo||(umOriginal&&!pesoUnitMap[umOriginal]&&!umCantidad&&!noAplica.peso))return null;
  const extra=Object.fromEntries(CAMPOS_EXTRA.map(([key])=>[key,noAplica[key]?["peso","totalMM","longitudIndividual"].includes(key)?null:"":["peso","totalMM","longitudIndividual"].includes(key)?numeroExcel(raw[key]):stringValue(key)]));
  if(!noAplica.totalMM&&extra.longitudIndividual!=null)extra.totalMM=cantidad*extra.longitudIndividual;
  return{codigo,descripcion,lote:extra.numeroLoteMP,proveedor:"",um:"PZ",cantidad,ubicacion,fechaRecibo,
    fechaModificacion:new Date().toLocaleString("es-MX"),tipoRegistro:esMP?"MP":esConsumible?"CONSUMIBLES":"GENERAL",seccionMaterial,esMP,peso,pesoU:extra.peso==null?null:(pesoUnitMap[umOriginal]||null),pesoUnidadPendiente:extra.peso!=null&&umCantidad,longitud:null,longU:null,
    seccion:seccionMaterial==="RND"?"redondo":["SQ","FLAT"].includes(seccionMaterial)?"cuadrado":null,secU:"mm",
    diametro:seccionMaterial==="RND"?numeroExcel(get("diametro")):null,
    alto:["SQ","FLAT"].includes(seccionMaterial)?numeroExcel(get("alto")):null,
    ancho:["SQ","FLAT"].includes(seccionMaterial)?numeroExcel(get("ancho")):null,...extra,noAplica,despachos:[]};
}
function consolidarExcelPorCodigo(materiales){
  const porCodigo=new Map();let filasAgrupadas=0;
  const camposIguales=["tipoRegistro","seccionMaterial","tipoMaterial","especificacionMedida","manufactura","totalMM","ubicacion","alto","ancho","diametro"];
  const camposCombinables=["numeroHeat","numeroLoteMP","numeroCertificado"];
  materiales.forEach(m=>{
    const clave=m.codigo.toLocaleLowerCase(),anterior=porCodigo.get(clave);
    if(!anterior){porCodigo.set(clave,m);return}
    for(const campo of camposIguales.filter(c=>c!=="totalMM")){
      const a=String(anterior[campo]??"").trim(),b=String(m[campo]??"").trim();
      if(a&&b&&a!==b)throw new Error(`El código ${m.codigo} aparece con valores distintos en "${campo}". Corrija esas filas antes de importar.`);
      if(!a&&b)anterior[campo]=m[campo];
    }
    if(anterior.longitudIndividual!=null&&m.longitudIndividual!=null&&anterior.longitudIndividual!==m.longitudIndividual)
      throw new Error(`El código ${m.codigo} aparece con longitudes individuales distintas. Corrija esas filas antes de importar.`);
    if(anterior.longitudIndividual==null)anterior.longitudIndividual=m.longitudIndividual;
    camposCombinables.forEach(campo=>{
      const valores=[...new Set([anterior[campo],m[campo]].flatMap(v=>String(v??"").split(" / ")).map(v=>v.trim()).filter(Boolean))];
      anterior[campo]=valores.join(" / ");
    });
    anterior.lote=anterior.numeroLoteMP;
    if(anterior.peso!=null&&m.peso!=null&&anterior.pesoU!==m.pesoU)
      throw new Error(`El código ${m.codigo} tiene pesos con unidades distintas o sin unidad. No se pueden sumar con seguridad.`);
    anterior.cantidad+=m.cantidad;
    if(m.peso!=null){
      anterior.peso=anterior.peso==null?m.peso:anterior.peso+m.peso;
      if(!anterior.pesoU)anterior.pesoU=m.pesoU;
    }
    anterior.pesoUnidadPendiente=anterior.peso!=null&&!anterior.pesoU;
    if(anterior.longitudIndividual!=null&&!anterior.noAplica?.totalMM)
      anterior.totalMM=anterior.cantidad*anterior.longitudIndividual;
    else if(anterior.totalMM!=null&&m.totalMM!=null)anterior.totalMM+=m.totalMM;
    const descripciones=[...new Set([anterior.descripcion,m.descripcion].filter(Boolean))];
    anterior.descripcion=descripciones.join(" / ");
    anterior.fechaRecibo=[anterior.fechaRecibo,m.fechaRecibo].sort()[0];
    const ordenes=[...new Set([anterior.ordenMaterial,m.ordenMaterial].flatMap(v=>String(v??"").split(" / ")).map(v=>v.trim()).filter(Boolean))];
    anterior.ordenMaterial=ordenes.join(" / ");
    filasAgrupadas++;
  });
  return{materiales:[...porCodigo.values()],filasAgrupadas};
}
async function importarExcel(file){
  if(!esAdmin)throw new Error("Solo el administrador puede importar información.");
  await cargarXLSX();
  const workbook=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true}),nombreHoja=workbook.SheetNames[0],sheet=workbook.Sheets[nombreHoja];
  if(!sheet)throw new Error("El archivo no contiene hojas para importar.");
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:true}),{index,mapped}=buscarEncabezados(rows);
  const esMateriaPrima=/materia[\s_-]*prima|\bmp\b/i.test(`${nombreHoja} ${file.name||""}`);
  let invalidos=0;const filas=[];
  rows.slice(index+1).forEach(row=>{
    if(!row.some(v=>String(v??"").trim()))return;
    const material=materialDesdeExcel(row,mapped,esMateriaPrima);if(!material){invalidos++;return}
    filas.push(material);
  });
  const{materiales:consolidados,filasAgrupadas}=consolidarExcelPorCodigo(filas);
  const codigosExistentes=new Set(mats.map(m=>m.codigo.toLocaleLowerCase()));
  const nuevos=consolidados.filter(m=>!codigosExistentes.has(m.codigo.toLocaleLowerCase()));
  const existentes=consolidados.length-nuevos.length;
  if(!nuevos.length){toast(`No se agregaron materiales. Ya existentes: ${existentes} · filas agrupadas: ${filasAgrupadas} · inválidas: ${invalidos}`,1);return}
  try{await subirLista(nuevos,[]);await cargar();renderHist();toast(`Importados: ${nuevos.length} · filas agrupadas: ${filasAgrupadas} · ya existentes: ${existentes} · inválidas: ${invalidos}`)}
  catch(err){errDB(err,"No se pudo importar el Excel");await cargar().catch(()=>{})}
}
$("fileExcel").onchange=async e=>{
  const file=e.target.files[0];e.target.value="";if(!file)return;
  try{await importarExcel(file)}catch(err){toast(err.message||"No se pudo leer el archivo de Excel",1)}
};
$("fileImp").onchange=e=>{
  if(!esAdmin){e.target.value="";toast("Solo el administrador puede restaurar información.",1);return}
  const f=e.target.files[0];if(!f)return;const r=new FileReader();
  r.onload=async()=>{try{
    const a=JSON.parse(r.result);if(!Array.isArray(a))throw 0;let nuevos=0,omit=0,surt=0;
    const antesC=new Set(mats.map(x=>x.codigo.toLowerCase())),antesT=new Set(mats.flatMap(x=>(x.despachos||[]).map(d=>x.codigo.toLowerCase()+"|"+d.t)));
    a.forEach(m=>{if(!m||!m.codigo){omit++;return}Object.assign(m,normalizarMaterial(m));m.codigo=String(m.codigo);m.esMP=!!m.esMP;m.cantidad=Number(m.cantidad)||0;m.despachos=Array.isArray(m.despachos)?m.despachos.filter(d=>d&&typeof d==="object"):[];const ex=mats.find(x=>x.codigo.toLowerCase()===String(m.codigo).toLowerCase());if(ex){const ids=new Set((ex.despachos||[]).map(d=>d.t));let add=0;(m.despachos||[]).forEach(d=>{if(!ids.has(d.t)){(ex.despachos=ex.despachos||[]).push(d);add++}});if(add)surt+=add;else omit++;return}mats.push(m);nuevos++});
    const msN=mats.filter(m=>!antesC.has(m.codigo.toLowerCase())),dsN=mats.flatMap(m=>(m.despachos||[]).filter(d=>d&&d.t!=null&&!antesT.has(m.codigo.toLowerCase()+"|"+d.t)).map(d=>({codigo:m.codigo,d})));
    try{await subirLista(msN,dsN)}catch(err){errDB(err,"No se pudo subir a Supabase");await cargar().catch(()=>{});return}
    await cargar().catch(()=>{});renderHist();toast(`Nuevos: ${nuevos} · surtidos fusionados: ${surt} · sin cambios: ${omit}`);
  }catch{toast("Archivo JSON no válido",1)}};
  r.readAsText(f);e.target.value="";
};

// --- Surtido ---
const FAC={mm:1,cm:10,m:1000,in:25.4};
const r4=n=>Math.round(n*1e4)/1e4;
function okU(u){return Object.prototype.hasOwnProperty.call(FAC,u)}
function longBase(m){
  const individual=Number(m.longitudIndividual);
  if(Number.isFinite(individual)&&individual>0)return{L:individual,u:"mm"};
  const L=Number(m.longitud),u0=m.longU||m.longitudUnidad;
  return{L,u:okU(u0)?u0:"mm"};
}
function modoLong(m){return m.esMP&&m.um==="PZ"&&longBase(m).L>0}
function longitudEnBase(d,u){
  const longitud=Number(d.longitud)||0,unidad=okU(d.unidad)?d.unidad:u;
  return longitud*FAC[unidad]/FAC[u];
}
function pesoUnidad(m){const u=m.pesoU||m.pesoUnidad;return m.peso!=null&&!m.pesoUnidadPendiente&&Number.isFinite(Number(m.peso))&&Number(m.peso)>=0&&["LB","KG"].includes(u)?u:""}
function pesoEquivalente(m,longitud,total){
  const u=pesoUnidad(m),peso=Number(m.peso);
  return u&&total>0&&Number.isFinite(peso)?{peso:r4(peso*longitud/total),u}:null;
}
function existenciaCalc(m){
  const ds=m.despachos||[];
  const ajustes=(m.ajustesInventario||[]).reduce((s,a)=>s+Number(a.diferencia||0),0);
  if(modoLong(m)){
    const{L,u}=longBase(m),total=m.cantidad*L;
    const surtido=ds.reduce((s,d)=>s+longitudEnBase(d,u),0),rest=r4(total-surtido+ajustes);
    return{modo:"long",u,L,total,rest,qty:r4(rest/L),peso:pesoEquivalente(m,rest,total)};
  }
  return{modo:"pz",u:m.um,rest:r4(m.cantidad-ds.reduce((s,d)=>s+(d.cantidad||0),0)+ajustes)}
}
function existTxt(m){const e=existencia(m);return e.modo==="long"?`${fmt(e.rest)} ${e.u} (≈${fmt(e.qty)} pz)${e.peso?` · ${fmt(e.peso.peso)} ${e.peso.u}`:""}`:`${fmt(e.rest)} ${e.u}`}
let surItems=[],nextSurItem=1,stream=null,scanning=false,modoEscaneo="surtido";

document.querySelectorAll(".tab").forEach(b=>b.onclick=()=>{
  document.querySelectorAll(".tab").forEach(x=>{x.classList.toggle("on",x===b);if(x===b)x.setAttribute("aria-current","page");else x.removeAttribute("aria-current")});
  const vista=b.dataset.v;
  $("vReg").hidden=vista!=="vReg";$("vSur").hidden=vista!=="vSur";$("vInv").hidden=vista!=="vInv";
  if(vista!=="vInv")limpiarConsultaUbicacion();
  if(vista==="vSur"){activarSubpantalla("vSur","preparar");modoEscaneo="surtido";renderSur();renderHist();enfocar()}else{camOff();if(vista==="vInv"){activarSubpantalla("vInv","existencias");modoEscaneo="inventario";renderInventario();renderUbicaciones();renderHistUbicaciones();renderAsignacion()}else{activarSubpantalla("vReg","captura");modoEscaneo="surtido";render()}}
});
const MAPA_SUBPANTALLAS={vInv:{existencias:["existencias"],almacenes:["almacenes","ubicaciones"],asignar:["asignar","historial"]}};
function activarSubpantalla(modulo,pagina){
  const vista=$(modulo);if(!vista)return;
  const visibles=MAPA_SUBPANTALLAS[modulo]?.[pagina]||[pagina];
  vista.querySelectorAll("[data-module-screen]").forEach(panel=>{panel.hidden=!visibles.includes(panel.dataset.moduleScreen)&&!(modulo==="vReg"&&(pagina==="captura"||pagina==="historial")&&panel.dataset.moduleScreen==="qr")});
  vista.querySelectorAll(".section-tab").forEach(b=>{if(b.dataset.page===pagina)b.setAttribute("aria-current","page");else b.removeAttribute("aria-current")});
  if(modulo==="vInv"){
    if(pagina==="almacenes"){renderAlmacenes();renderUbicaciones();}
    if(pagina==="asignar"){renderAsignacion();renderHistUbicaciones();}
    if(pagina==="existencias"){renderInventario();}
    if(pagina!=="asignar"&&modoEscaneo==="inventario")camOff();
  }
  if(modulo==="vSur"&&pagina!=="preparar")camOff();
}
document.querySelectorAll(".section-tab").forEach(b=>b.onclick=()=>activarSubpantalla(b.dataset.module,b.dataset.page));
function enfocar(){if(matchMedia("(pointer:fine)").matches)$("scanIn").focus()}

function escanear(txt){
  if(modoEscaneo==="ubiConsulta"){procesarEscaneoUbiConsulta(txt);return}
  if(modoEscaneo==="inventario"){procesarEscaneoInventario(txt);return}
  const cod=codDe(txt).toLowerCase();
  if(modoEscaneo==="registro"){
    camOff();
    const material=mats.find(m=>m.codigo.toLowerCase()===cod);
    if(!material){toast(cod?`Código no registrado: ${cod}`:"El QR no contiene un código de material válido.",1);return}
    mostrar(material);
    $("vReg").querySelector('[data-module-screen="qr"]').scrollIntoView({behavior:"smooth",block:"start"});
    toast(`Material encontrado: ${material.codigo}`);return
  }
  const i=mats.findIndex(m=>m.codigo.toLowerCase()===cod);
  if(i<0){toast(`Código no registrado: ${cod}`,1);return}
  const pz=pzDe(txt);if(pz){escanearPieza(i,pz.n);return}
  agregarSurItem(i);renderSur();
}
// Los QR de pieza se agregan como renglones individuales y no permiten duplicados.
function escanearPieza(i,n){
  const m=mats[i];
  if(!esIndividual(m)||n<1||n>m.cantidad){toast(`Pieza ${n} no válida para ${m.codigo}`,1);return}
  if(piezasSurt(m).has(n)){toast(`⚠ La pieza ${n}/${m.cantidad} de ${m.codigo} ya fue surtida`,1);return}
  if(surItems.some(x=>x.codigo.toLowerCase()===m.codigo.toLowerCase()&&x.piezas?.includes(n))){toast(`La pieza ${n} ya está en el surtido actual`,1);return}
  agregarSurItem(i,n);renderSur();
}
function agregarSurItem(i,pieza=null){
  const m=mats[i],e=existencia(m),longitud=pieza&&e.modo==="long"?String(e.L):"";
  surItems.push({id:nextSurItem++,codigo:m.codigo,cantidad:1,longitud,unidad:e.modo==="long"?e.u:"",piezas:pieza?[pieza]:[]});
}
const sc=$("scanIn");let sT,ubiScanT;
sc.addEventListener("input",()=>{clearTimeout(sT);if(/Código:/.test(sc.value)||/#[dc]=/.test(sc.value)||(modoEscaneo==="ubiConsulta"&&/^TRAZABILIDAD/i.test(sc.value)))sT=setTimeout(()=>{if(sc.value){escanear(sc.value);sc.value=""}},250)});
sc.addEventListener("keydown",e=>{
  if(e.key!=="Enter"||e.shiftKey)return;const v=sc.value.trim();
  if(!v||(/^TRAZABILIDAD/i.test(v)&&modoEscaneo!=="ubiConsulta"))return;
  e.preventDefault();clearTimeout(sT);escanear(v);sc.value="";
});
$("ubiScan").addEventListener("input",()=>{clearTimeout(ubiScanT);if($("ubiScan").value.includes("TRAZABILIDAD-UBICACION:")||/#[dc]=/.test($("ubiScan").value))ubiScanT=setTimeout(manejarLectorUbicacion,250)});

// Cámara (BarcodeDetector si existe; si no, jsQR)
let det=null,audioCtx=null,wake=null,torchOn=false,busy=false,tUlt=0,txtUlt="",vistoUlt=0;
try{if("BarcodeDetector" in window)det=new BarcodeDetector({formats:["qr_code"]})}catch{}
function compacto(){try{return localStorage.getItem("qr_compacto2")!=="0"}catch{return true}}
$("qrCompacto").checked=compacto();
$("qrCompacto").onchange=()=>{try{localStorage.setItem("qr_compacto2",$("qrCompacto").checked?"1":"0")}catch{}if(actual)mostrar(actual)};
if(!window.isSecureContext||!navigator.mediaDevices){$("camAviso").hidden=false;$("invCamAviso").hidden=false;$("regCamAviso").hidden=false}
if(matchMedia("(pointer:fine)").matches)$("manual").open=true;
function feedback(){
  navigator.vibrate&&navigator.vibrate(60);
  if(audioCtx){const o=audioCtx.createOscillator(),g=audioCtx.createGain();o.frequency.value=880;g.gain.value=.15;o.connect(g);g.connect(audioCtx.destination);o.start();o.stop(audioCtx.currentTime+.09)}
  const w=$("camWrap");w.classList.add("ok");setTimeout(()=>w.classList.remove("ok"),300);
}
async function camOn(){
  if(!window.isSecureContext||!navigator.mediaDevices){$("camAviso").hidden=false;$("invCamAviso").hidden=false;$("regCamAviso").hidden=false;toast("La cámara en vivo requiere HTTPS",1);return}
  try{
    audioCtx=audioCtx||new(window.AudioContext||window.webkitAudioContext)();
    stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"},width:{ideal:1920},height:{ideal:1080}},audio:false});
    const tr=stream.getVideoTracks()[0],cap=tr.getCapabilities?tr.getCapabilities():{};
    if(cap.focusMode&&cap.focusMode.includes("continuous"))tr.applyConstraints({advanced:[{focusMode:"continuous"}]}).catch(()=>{});
    torchOn=false;$("bTorch").hidden=!cap.torch;
    $("video").srcObject=stream;await $("video").play();$("camWrap").hidden=false;scanning=true;$("bCam").textContent="⏹ Detener cámara";$("bCamInv").textContent="⏹ Detener cámara";$("bCamReg").textContent="⏹ Detener cámara";
    try{wake=await navigator.wakeLock.request("screen")}catch{}
    bucle();
  }catch(e){toast(e.name==="NotAllowedError"?"Permiso de cámara denegado. Actívelo en los ajustes del navegador.":"No se pudo abrir la cámara.",1)}
}
function camOff(){scanning=false;if(stream)stream.getTracks().forEach(t=>t.stop());stream=null;if(wake){wake.release().catch(()=>{});wake=null}$("camWrap").hidden=true;$("bCam").textContent="📷 Escanear con cámara";$("bCamInv").textContent="📷 Escanear ubicación y materiales";$("bCamReg").textContent="📷 Escanear código y ver información"}
$("bCam").onclick=()=>{if(scanning){camOff();return}modoEscaneo="surtido";camOn()};
$("bCamInv").onclick=()=>{if(scanning){camOff();return}modoEscaneo="inventario";camOn()};
$("bCamReg").onclick=()=>{if(scanning){camOff();return}modoEscaneo="registro";camOn()};
$("bCamClose").onclick=camOff;
$("bTorch").onclick=()=>{torchOn=!torchOn;stream.getVideoTracks()[0].applyConstraints({advanced:[{torch:torchOn}]}).catch(()=>toast("Linterna no disponible",1))};
document.addEventListener("visibilitychange",()=>{if(document.hidden)camOff()});
async function leer(v){if(!det&&!window.jsQR)await cargarJsQR();
  if(det){try{const r=await det.detect(v);return r[0]?r[0].rawValue:null}catch{det=null}}
  const w=Math.min(v.videoWidth,1100),hh=Math.round(v.videoHeight*w/v.videoWidth),cv=$("camCv");
  cv.width=w;cv.height=hh;const x=cv.getContext("2d",{willReadFrequently:true});x.drawImage(v,0,0,w,hh);
  const r=jsQR(x.getImageData(0,0,w,hh).data,w,hh,{inversionAttempts:"dontInvert"});return r&&r.data}
function bucle(){
  if(!scanning)return;const v=$("video"),now=Date.now();
  if(!busy&&v.readyState===4&&now-tUlt>120){busy=true;tUlt=now;
    leer(v).then(t=>{const n=Date.now();if(t){vistoUlt=n;if(t!==txtUlt){txtUlt=t;feedback();escanear(t)}}else if(n-vistoUlt>700)txtUlt=""}).catch(()=>{}).finally(()=>{busy=false})}
  requestAnimationFrame(bucle);
}
// Panel de surtido
function renderSur(){
  const P=$("sPanel");
  if(!surItems.length){P.innerHTML='<div class="empty">Escanee los QR de los materiales para agregarlos a la salida</div>';return}
  let h="";
  h+=surItems.map((item,n)=>{
    const m=mats.find(x=>x.codigo.toLowerCase()===item.codigo.toLowerCase());if(!m)return"";
    const e=existencia(m),lote=m.numeroLoteMP||m.lote||"";
    return`<section class="mp"><div class="sur-item-head"><b>${n+1}. ${esc(m.codigo)} · ${esc(m.descripcion)}</b><button type="button" class="bd" data-quitar="${item.id}">Quitar</button></div>
      ${lote||m.numeroHeat?`<small class="sur-item-meta">Lote / Heat: ${esc([lote,m.numeroHeat].filter(Boolean).join(" / "))}</small>`:""}
      <div class="sur-item-qty"><span><b>Existencia actual:</b> ${esc(existTxt(m))}</span>
      ${e.modo==="long"?`<div class="sur-entry"><label for="sLen-${item.id}">Longitud (pieza de ${esc(fmt(e.L))} ${esc(e.u)})</label><input type="number" id="sLen-${item.id}" min="0" step="any" inputmode="decimal" placeholder="Longitud" value="${esc(item.longitud)}"><select id="sU-${item.id}">${["mm","cm","m","in"].map(u=>`<option value="${u}"${(item.unidad||e.u)===u?" selected":""}>${u}</option>`).join("")}</select></div>`:
      `<label class="sur-entry"><span>Cantidad (${esc(m.um)})</span><input type="number" id="sCnt-${item.id}" min="0" step="any" value="${esc(item.cantidad)}"></label>`}</div>
      ${item.piezas?.length?`<small class="sur-item-meta">Pieza QR: ${item.piezas.join(", ")}</small>`:""}
      <div class="res" id="sItemRes-${item.id}"></div></section>`;
  }).join("");
  h+=`<div class="res" id="sRes"></div><div class="acts sur-actions"><button type="button" class="bp" id="sOk">✅ Confirmar salida múltiple</button><button type="button" class="bd" id="sNo">Vaciar salida</button></div>`;
  P.innerHTML=h;
  surItems.forEach(item=>{
    const len=$(`sLen-${item.id}`),unit=$(`sU-${item.id}`),cnt=$(`sCnt-${item.id}`);
    if(len)len.oninput=()=>{item.longitud=len.value;calcSur()};
    if(unit)unit.onchange=()=>{item.unidad=unit.value;calcSur()};
    if(cnt)cnt.oninput=()=>{item.cantidad=Number(cnt.value);calcSur()};
  });
  P.querySelectorAll("[data-quitar]").forEach(b=>b.onclick=()=>{surItems=surItems.filter(x=>x.id!==Number(b.dataset.quitar));renderSur()});
  $("sNo").onclick=()=>{surItems=[];renderSur();enfocar()};$("sOk").onclick=confirmarSur;calcSur();
}
function calcSur(){
  const totalPorCodigo=new Map(),salidas=[],errores=[];
  surItems.forEach(item=>{
    const m=mats.find(x=>x.codigo.toLowerCase()===item.codigo.toLowerCase());
    if(!m){errores.push(`${esc(item.codigo)} ya no está registrado`);return}
    const e=existencia(m),anterior=totalPorCodigo.get(m.codigo.toLowerCase())||0;
    if(item.errorPieza){errores.push(`${esc(m.codigo)}: una pieza seleccionada ya fue surtida`);return}
    let d=null,html="";
    if(e.modo==="long"){
      const len=Number(item.longitud)||0,u=okU(item.unidad)?item.unidad:e.u;
      const lr=r4(len*FAC[u]/FAC[e.u]),disponible=r4(e.rest-anterior),quedan=r4(disponible-lr);
      if(len<=0){html="Capture la longitud a surtir."}
      else if(quedan<0){html=`⚠ Excede la existencia disponible para este renglón (${fmt(disponible)} ${esc(e.u)}).`;errores.push(`${esc(m.codigo)}: longitud total excedida`)}
      else{
        const peso=pesoEquivalente(m,lr,e.total),pesoRest=pesoEquivalente(m,quedan,e.total);
        const piezas=item.piezas?.length===r4(lr/e.L)?item.piezas.slice():undefined;
        d={modo:"longitud",longitud:lr,unidad:e.u,cantidad:r4(lr/e.L),piezasEq:r4(lr/e.L),...(piezas?{piezas}:{}),...(peso?{pesoSurtido:peso.peso,pesoU:peso.u}:{})};
        html=`Surtir: <b>${fmt(lr)} ${esc(e.u)}</b>${peso?` · ≈ ${fmt(peso.peso)} ${peso.u}`:""}<br>Restante después de este renglón: ${fmt(quedan)} ${esc(e.u)}${pesoRest?` · ≈ ${fmt(pesoRest.peso)} ${pesoRest.u}`:""}`;
        totalPorCodigo.set(m.codigo.toLowerCase(),anterior+lr);
      }
    }else{
      const qty=Number(item.cantidad)||0,disponible=r4(e.rest-anterior),quedan=r4(disponible-qty);
      if(qty<=0){html="Indique una cantidad mayor que cero."}
      else if(quedan<0){html=`⚠ Excede la existencia disponible para este renglón (${fmt(disponible)} ${esc(e.u)}).`;errores.push(`${esc(m.codigo)}: cantidad total excedida`)}
      else{
        d={modo:"pieza",cantidad:qty,...(item.piezas?.length===qty?{piezas:item.piezas.slice()}:{})};
        html=`Surtir: <b>${fmt(qty)} ${esc(e.u)}</b><br>Restante después de este renglón: ${fmt(quedan)} ${esc(e.u)}`;
        totalPorCodigo.set(m.codigo.toLowerCase(),anterior+qty);
      }
    }
    const R=$(`sItemRes-${item.id}`);if(R){R.innerHTML=html;R.classList.toggle("bad",!d)}
    item.d=d;
    if(d)salidas.push({item,m,d});
  });
  const R=$("sRes");
  if(R){R.innerHTML=errores.length?`⚠ Revise antes de confirmar:<br>${errores.map(x=>`• ${x}`).join("<br>")}`:`${salidas.length} renglón(es) listos para registrar en una salida.`;R.classList.toggle("bad",errores.length>0)}
  $("sOk").disabled=!salidas.length||errores.length>0||salidas.length!==surItems.length;
  calcSur.items=errores.length?null:salidas;
}
async function confirmarSur(){
  const lineas=calcSur.items;if(!lineas?.length)return;const orden=$("orden").value.trim();
  if(!orden){toast("Capture la orden de producción para generar el comprobante de salida.",1);$("orden").focus();return}
  const t=Date.now(),folio=folioSalida(t),fecha=new Date(t).toLocaleString("es-MX"),porCodigo=new Map(),registros=[];
  lineas.forEach(({m,d},i)=>{
    const anteriores=porCodigo.get(m.codigo.toLowerCase())||[];
    const temporal={...m,despachos:[...(m.despachos||[]),...anteriores]};
    const reg={...d,t:t+i,folio,fecha,orden,
      entrega:$("entrega").value.trim(),recibe:$("recibe").value.trim(),maquina:$("maquina").value.trim(),
      motivo:$("motivo").value.trim(),observaciones:$("observaciones").value.trim(),
      material:{codigo:m.codigo,descripcion:m.descripcion,tipoRegistro:m.tipoRegistro,lote:m.numeroLoteMP||m.lote,numeroHeat:m.numeroHeat,ubicacion:m.ubicacion,seccionMaterial:m.seccionMaterial,especificacionMedida:m.especificacionMedida}};
    const exPost=existencia({...temporal,despachos:[...temporal.despachos,reg]});
    reg.existenciaRestante=exPost.rest;reg.existenciaUnidad=exPost.u;
    if(exPost.peso){reg.pesoRestante=exPost.peso.peso;reg.pesoRestanteU=exPost.peso.u}
    porCodigo.set(m.codigo.toLowerCase(),[...anteriores,reg]);registros.push({m,d:reg});
  });
  $("sOk").disabled=true;
  try{
    const filas=registros.map(({m,d})=>desFila(m.codigo,d)),{error}=await sb.from("despachos").insert(filas);
    if(error)throw error;
  }catch(err){errDB(err,"No se pudo registrar la salida múltiple");await cargar().catch(()=>{});renderSur();return}
  await cargar().catch(()=>{});
  const confirmados=registros.map(({m,d})=>({m:mats.find(x=>x.codigo.toLowerCase()===m.codigo.toLowerCase())||m,d}));
  imprimirSalidaGrupo(confirmados,folio);
  toast(`Salida registrada ✔ · ${registros.length} renglón(es) · Folio ${folio}`);
  surItems=[];renderSur();renderHist();enfocar();
}
function folioSalida(t){
  const fecha=new Date(t),dia=`${fecha.getFullYear()}${String(fecha.getMonth()+1).padStart(2,"0")}${String(fecha.getDate()).padStart(2,"0")}`;
  return`SAL-${dia}-${t}`;
}
function movs(){const L=[];mats.forEach((m,i)=>(m.despachos||[]).forEach((d,j)=>L.push({m,i,j,d})));return L.sort((a,b)=>b.d.t-a.d.t)}
function pesoSurtido(m,d){
  const u=d.pesoU||pesoUnidad(m);
  if(d.pesoSurtido!=null&&["LB","KG"].includes(u))return{peso:Number(d.pesoSurtido),u};
  if(d.modo==="longitud"){
    const e=existencia(m);
    return pesoEquivalente(m,longitudEnBase(d,e.u),e.total);
  }
  return null;
}
function surTxt(m,d){
  const base=d.modo==="longitud"?`${fmt(d.longitud)} ${d.unidad} ≈ ${fmt(d.piezasEq)} pz`:`${fmt(d.cantidad)} ${m.um}`;
  const peso=d.modo==="longitud"?pesoSurtido(m,d):null;
  return base+(peso?` · ≈ ${fmt(peso.peso)} ${peso.u}`:"")+(d.piezas&&d.piezas.length?` (pzas ${d.piezas.join(",")})`:"");
}
function imprimirSalida(m,d){
  imprimirSalidaGrupo([{m,d}],d.folio||folioSalida(d.t));
}
function imprimirSalidaGrupo(items,folio){
  if(!items.length)return;
  const {d:base}=items[0];
  const encabezados=[
    ["Fecha y hora",base.fecha],["Orden de producción",base.orden],
    ["Máquina / línea",base.maquina],["Motivo / uso",base.motivo]
  ];
  const renglones=items.map(({m,d})=>{
    const p=d.modo==="longitud"?pesoSurtido(m,d):null;
    const surtido=d.modo==="longitud"?`${fmt(d.longitud)} ${d.unidad} (${fmt(d.piezasEq)} pz eq.)`:`${fmt(d.cantidad)} ${m.um}`;
    const despues={...m,despachos:[...(m.despachos||[]).filter(x=>x.t!==d.t),d]},rest=existencia(despues);
    const restL=d.existenciaRestante??rest.rest,restU=d.existenciaUnidad||rest.u;
    const restPeso=d.pesoRestante!=null?{peso:d.pesoRestante,u:d.pesoRestanteU}:rest.peso;
    const restantes=`${fmt(restL)} ${restU}${restPeso?` (≈ ${fmt(restPeso.peso)} ${restPeso.u})`:""}`;
    const material=d.material||m;
    return`<tr><td>${esc(material.codigo)}</td><td>${esc(nombreTipoRegistro(material.tipoRegistro||(m.esMP?"MP":"GENERAL")))}</td><td>${esc(material.descripcion)}</td>
      <td>${esc([material.lote,material.numeroHeat].filter(Boolean).join(" / ")||"—")}</td>
      <td>${esc([material.seccionMaterial,material.especificacionMedida].filter(Boolean).join(" · ")||"—")}</td>
      <td>${esc(material.ubicacion||"—")}</td><td>${esc(surtido)}</td>
      <td>${esc(p?`${fmt(p.peso)} ${p.u}`:"—")}</td><td>${esc(restantes)}</td></tr>`;
  }).join("");
  $("salidaView").innerHTML=`<article class="salida">
    <div class="salida-top"><div><h1>Comprobante de salida de materiales</h1><div>Trazabilidad de materiales · Surtido a Producción</div></div><div class="salida-meta"><div class="doc-code">Doc-N-FAL-0102R0</div><div class="folio">Folio ${esc(folio)}</div></div></div>
    <h2>Datos de la salida</h2><table class="datos"><tbody>${encabezados.map(([k,v])=>`<tr><th>${esc(k)}</th><td>${esc(v||"—")}</td></tr>`).join("")}
      <tr><th>Entrega · Almacén</th><td>${esc(base.entrega||"—")}</td></tr><tr><th>Recibe · Producción</th><td>${esc(base.recibe||"—")}</td></tr></tbody></table>
    <h2>Materiales surtidos · ${items.length} renglón(es)</h2><table class="renglones"><colgroup><col style="width:10%"><col style="width:10%"><col style="width:21%"><col style="width:10%"><col style="width:12%"><col style="width:8%"><col style="width:9%"><col style="width:8%"><col style="width:12%"></colgroup><thead><tr><th>Código</th><th>Clasificación</th><th>Descripción</th><th>Lote /<br>Heat</th><th>Sección /<br>especificación</th><th>Ubicación</th><th>Surtido</th><th>Peso eq.</th><th>Existencia<br>restante</th></tr></thead><tbody>${renglones}</tbody></table>
    <h2>Observaciones</h2><div class="salida-nota">${esc(base.observaciones||"")}</div>
    <div class="firmas"><div><div class="firma-linea"></div><div>Firma de quien entrega · Almacén ${esc(base.entrega||"")}</div></div>
    <div><div class="firma-linea"></div><div>Firma de quien recibe · Producción ${esc(base.recibe||"")}</div></div></div>
    <p class="salida-pie">Conserve este comprobante como evidencia de entrega. El folio identifica todos los movimientos agrupados en esta salida.</p>
  </article>`;
  $("salidaView").hidden=false;
  const estiloImpresion=document.createElement("style");
  estiloImpresion.textContent="@page salida{size:A4 portrait;margin:8mm}";
  document.head.append(estiloImpresion);
  const tituloAnterior=document.title;
  document.title=folio;
  try{window.print()}
  finally{$("salidaView").hidden=true;estiloImpresion.remove();document.title=tituloAnterior}
}
function imprimirFolio(folio){
  const items=movs().filter(({d})=>(d.folio||folioSalida(d.t))===folio);
  if(items.length)imprimirSalidaGrupo(items,folio);
}
function renderHist(){
  const L=movs().slice(0,100),groups=new Map();
  L.forEach(row=>{
    const order=String(row.d.orden||"").trim(),key=order?`orden:${order.toLocaleLowerCase()}`:`sin-orden:${row.m.codigo}:${row.d.t}`;
    if(!groups.has(key))groups.set(key,{orden:order,items:[]});
    groups.get(key).items.push(row);
  });
  $("hBody").innerHTML=[...groups.values()].map(g=>{
    const materiales=new Set(g.items.map(({m})=>m.codigo)),ultima=g.items[0].d.fecha||"";
    const detalles=g.items.map(({m,i,j,d})=>`<tr><td>${esc(d.fecha)}</td><td><b>${esc(m.codigo)}</b><br><small>${esc(d.folio||folioSalida(d.t))}</small></td><td>${esc(surTxt(m,d))}</td><td><button type="button" class="ic" data-action="pdf" data-folio="${esc(d.folio||folioSalida(d.t))}" title="PDF / Imprimir salida completa">📄</button>${esAdmin?` <button type="button" class="ic" data-action="anular" data-i="${i}" data-j="${j}" title="Anular este renglón">↩️</button>`:""}</td></tr>`).join("");
    return`<details class="sur-group"><summary><span class="sur-date">${esc(ultima)}</span><span class="sur-order">Orden: ${esc(g.orden||"Sin orden")}</span><span class="sur-count">${g.items.length} surtido(s) · ${materiales.size} material(es)</span></summary>
      <div class="sur-group-body"><table><thead><tr><th>Fecha</th><th>Código / folio</th><th>Surtido</th><th>Acciones</th></tr></thead><tbody>${detalles}</tbody></table></div></details>`;
  }).join("");
  $("hVacio").hidden=L.length>0;
}
$("hBody").addEventListener("click",async e=>{
  const b=e.target.closest("button");if(!b)return;
  if(b.dataset.action==="pdf"){imprimirFolio(b.dataset.folio);return}
  if(!esAdmin){toast("Solo el administrador puede anular surtidos.",1);return}
  const m=mats[+b.dataset.i],d=m.despachos[+b.dataset.j];
  if(await confirmar(`¿Anular el surtido de ${m.codigo} (${surTxt(m,d)})? La existencia se restituye.`)){try{const{error}=await sb.from("despachos").delete().eq("codigo",m.codigo).eq("t",d.t);if(error)throw error;await cargar();renderHist();renderSur();toast("Surtido anulado")}catch(err){errDB(err,"No se pudo anular")}}
});
$("bCsvS").onclick=()=>{
  if(!esAdmin){toast("Solo el administrador puede exportar información.",1);return}
  const c=v=>{let t=String(v??"");if(/^[=+\-@\t\r]/.test(t))t="'"+t;return `"${t.replace(/"/g,'""')}"`};
  const f=movs().map(({m,d})=>{const p=d.modo==="longitud"?pesoSurtido(m,d):null;return[d.folio||folioSalida(d.t),d.fecha,d.orden,d.entrega,d.recibe,d.maquina,d.motivo,d.observaciones,m.codigo,m.descripcion,d.modo,d.cantidad,m.um,d.longitud??"",d.unidad??"",p?.peso??"",p?.u??"",(d.piezas||[]).join(" ")].map(c).join(",")});
  bajar("surtidos.csv","\ufeff"+["folio,fecha,orden,entrega,recibe,maquina,motivo,observaciones,codigo,descripcion,modo,cantidad,um,longitud,unidad,peso_equivalente,peso_um,piezas",...f].join("\r\n"),"text/csv;charset=utf-8");toast("CSV exportado ⬇️");
};

// --- Ficha: página que se abre al escanear el QR con la cámara del celular ---
function limpiaP(o){
  const t=v=>v==null?"":String(v).slice(0,300),n=v=>v==null?null:(Number(v)||0);
  return{codigo:t(o.c),descripcion:t(o.d),lote:t(o.l),proveedor:t(o.p),um:"PZ",cantidad:n(o.q)??0,esMP:!!o.m,ubicacion:t(o.ub),fechaRecibo:t(o.fr),fechaModificacion:t(o.fm),
    tipoRegistro:o.tr==="MP"||o.m?"MP":o.tr==="CONSUMIBLES"?"CONSUMIBLES":"GENERAL",seccionMaterial:["SQ","FLAT","RND"].includes(o.sm)?o.sm:"",tipoMaterial:t(o.tm),especificacionMedida:t(o.em),longitudIndividual:n(o.li),manufactura:t(o.mf),numeroHeat:t(o.nh),numeroLoteMP:t(o.nl),numeroCertificado:t(o.nc),totalMM:n(o.tmm),tipoMoneda:t(o.mon),ordenMaterial:t(o.ord),poNumber:t(o.po),noAplica:o.na&&typeof o.na==="object"?o.na:{},
    peso:n(o.pe),pesoU:o.pu==="LB"||o.pu==="KG"?o.pu:o.u==="LB"||o.u==="KG"?o.u:null,pesoUnidadPendiente:o.pu==="?",longitud:n(o.lo),longU:okU(o.lu)?o.lu:"mm",seccion:o.s==="cuadrado"?"cuadrado":"redondo",
    secU:["mm","cm","in"].includes(o.su)?o.su:"mm",diametro:n(o.di),alto:n(o.al),ancho:n(o.an),despachos:[],existSnap:t(o.x)}}
function filasFicha(m,ex,pz){
  const et=pz?`Pieza ${pz.n} de ${pz.t||m.cantidad}`:`MASTER · ${m.cantidad} ${m.um}`;
  const st=pz&&ex?(piezasSurt(m).has(pz.n)?"⚠ SURTIDA":"En existencia"):"";
  return[["Código",m.codigo],["Etiqueta",et],["Estado de la pieza",st],["Clasificación",nombreTipoRegistro(m.tipoRegistro)],["Tipo material",m.tipoMaterial],["Descripción",m.descripcion],["Proveedor",m.proveedor],
    ["Sección / especificación",seccionEspecificacion(m)],["Longitud IND",m.longitudIndividual!=null?`${m.longitudIndividual} mm`:""],["Manufactura",m.manufactura],["N.º Heat",m.numeroHeat],["N.º lote MP",m.numeroLoteMP||m.lote],["N.º cert.",m.numeroCertificado],["Total mm",m.totalMM],["Fecha de recibo",m.fechaRecibo],["Cantidad",`${m.cantidad} ${m.um}`],
    ["Peso",m.peso!=null?`${m.peso}${m.pesoU||m.pesoUnidad?" "+(m.pesoU||m.pesoUnidad):m.pesoUnidadPendiente?" (unidad sin definir)":""}`:""],["Orden",m.ordenMaterial],["PO number",m.poNumber],
    ["Existencia",ex?existTxt(m):""],["Existencia al imprimir la etiqueta",!ex&&m.existSnap?m.existSnap:""],["Últ. modificación",m.fechaModificacion]]}
function salirFicha(){document.body.classList.remove("ficha");$("fichaView").hidden=true;document.title="Trazabilidad de Material · QR";try{history.replaceState(null,"",location.pathname+location.search)}catch{}}
async function iniciarFicha(){
  const x=/#(?:[^#]*&)?(d|c)=([^&\s]+)/.exec(location.hash);if(!x)return;
  const pz=pzDe(location.hash);let snap=null,cod="";
  try{if(x[1]==="d"){snap=limpiaP(JSON.parse(b64d(x[2])));cod=snap.codigo}else cod=decodeURIComponent(x[2])}catch{}
  document.body.classList.add("ficha");$("fichaView").hidden=false;window.scrollTo(0,0);
  $("fApp").onclick=salirFicha;$("fPdf").onclick=()=>{limpiarEtqCss();window.print()};
  if(!cod){$("fCod").textContent="QR no válido";$("fDatos").innerHTML='<div class="res bad">No se pudo leer la información de este QR. Genere la etiqueta de nuevo.</div>';$("fNota").textContent="";$("fPdf").hidden=true;return}
  $("fCod").textContent=cod;document.title="Ficha "+cod;$("fDatos").innerHTML='<div class="res">Consultando…</div>';$("fNota").textContent="";$("fPdf").hidden=true;$("fSur").hidden=true;
  let m=mats.find(k=>String(k.codigo).toLowerCase()===cod.toLowerCase())||null,live=!!m,red=false,noReg=false;
  if(!m){
    try{if(!sb)throw 0;const r=await sb.rpc("ficha_publica",{p_codigo:cod});if(r.error)throw r.error;
      if(r.data){
        const a=await sb.rpc("ficha_existencia_publica",{p_codigo:cod});if(a.error)throw a.error;
        const diferencia=Number(a.data);if(a.data==null||!Number.isFinite(diferencia))throw new Error("La respuesta del ajuste de inventario no es válida.");
        m={...r.data.m,despachos:(r.data.d||[]).map(desDeFila),ajustesInventario:[{diferencia}]};live=true
      }else noReg=true}
    catch{red=true}
  }
  let solo=false;if(!m){if(snap)m=snap;else{m={codigo:cod};solo=true}}
  const hora=new Date().toLocaleString("es-MX");
  if(solo){
    $("fDatos").innerHTML=`<div class="res bad">${red?"No se pudo consultar el sistema. Revise su conexión a internet e intente de nuevo.":`El código <b>${esc(cod)}</b> no está registrado en el sistema.`}</div>`;
    $("fNota").textContent="";return}
  $("fPdf").hidden=false;
  $("fDatos").innerHTML=filasFicha(m,live,pz).filter(r=>r[1]).map(r=>`<div class="d"><b>${r[0]}</b><span>${esc(r[1])}</span></div>`).join("");
  $("fNota").textContent=live?`Datos en línea del sistema de trazabilidad. Consultado: ${hora}.`:`${noReg?"Este código ya no está registrado en el sistema.":"No se pudo consultar el sistema."} Se muestran los datos grabados al generar la etiqueta${m.fechaModificacion?" (últ. modificación: "+m.fechaModificacion+")":""}, que no reflejan surtidos posteriores. Consultado: ${hora}.`;
  $("fSur").hidden=!(esStaff&&live);
  $("fSur").onclick=()=>{const c=pz?location.hash:m.codigo;salirFicha();document.querySelector('.tab[data-v="vSur"]').click();escanear(c)};
}
window.addEventListener("hashchange",()=>{if(/#(?:[^#]*&)?[dc]=/.test(location.hash))location.reload()});

resetForm();render();arrancar();
