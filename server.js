// Quiz local type Kahoot - serveur Node.js + Socket.IO
// Tourne sur ton ordinateur. Les telephones se connectent via le meme WiFi.
const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const os = require('os');
const path = require('path');
const fs = require('fs');
const QRCode = require('qrcode');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3000;
const LIMIT = 20; // secondes par question (modifiable)

const app = express();
const server = http.createServer(app);
const io = new Server(server);
app.use(express.static(path.join(__dirname, 'public')));

let QUESTIONS = [];
let QUIZ_TITLE = 'Quiz';

// ---------- Analyse tolérante des fichiers ----------
function norm(s){ return String(s==null?'':s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,''); }
function letterToIndex(v){
  const s=String(v==null?'':v).trim().toUpperCase();
  if(s.length===1 && 'ABCDE'.includes(s)) return 'ABCDE'.indexOf(s);
  const n=parseInt(s,10); if(n>=1&&n<=5) return n-1;
  return -1;
}
function normImg(img){
  img=String(img==null?'':img).trim();
  if(!img) return '';
  if(/^https?:/i.test(img)) return img;
  if(img.indexOf('/')>=0) return '/'+img.replace(/^\/+/,'');
  return '/images/'+img;
}
function rowsToQuestions(rows){
  const out=[];
  rows.forEach(r=>{
    const m={}; Object.keys(r).forEach(k=>m[norm(k)]=r[k]);
    const q=m['question']||m['questiontext']||m['enonce']||m['q']||'';
    const o=[
      m['optiona']||m['answer1']||m['option1']||m['a']||m['reponse1']||m['choix1']||'',
      m['optionb']||m['answer2']||m['option2']||m['b']||m['reponse2']||m['choix2']||'',
      m['optionc']||m['answer3']||m['option3']||m['c']||m['reponse3']||m['choix3']||'',
      m['optiond']||m['answer4']||m['option4']||m['d']||m['reponse4']||m['choix4']||''
    ].map(x=>String(x==null?'':x).trim());
    const corr=m['bonnereponse']||m['correctanswers']||m['correctanswer']||m['correct']||m['reponsecorrecte']||m['bonne']||m['solution']||m['reponsen']||'';
    let a=letterToIndex(corr);
    if(a<0 && corr){ const ci=o.findIndex(x=>x && norm(x)===norm(corr)); if(ci>=0)a=ci; }
    const img=normImg(m['image']||m['img']||m['imagelink']||m['imagefacultatif']||'');
    const opts=o.filter(x=>x!=='');
    if(q && opts.length>=2 && a>=0 && a<opts.length){
      out.push({ q:String(q).trim(), o:opts, a, img });
    }
  });
  return out;
}
function workbookToQuiz(wb){
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
  return rowsToQuestions(rows);
}

// ---------- Chargement au démarrage (xlsx > csv > json) ----------
function loadFromDisk(){
  try{
    const xlsxP=path.join(__dirname,'questions.xlsx');
    const csvP =path.join(__dirname,'questions.csv');
    const jsonP=path.join(__dirname,'questions.json');
    if(fs.existsSync(xlsxP)){ return { title:'Quiz', questions:workbookToQuiz(XLSX.readFile(xlsxP)) }; }
    if(fs.existsSync(csvP)){ return { title:'Quiz', questions:workbookToQuiz(XLSX.read(fs.readFileSync(csvP,'utf8'),{type:'string'})) }; }
    if(fs.existsSync(jsonP)){
      const d=JSON.parse(fs.readFileSync(jsonP,'utf8'));
      const questions=(d.questions||d.q||[]).map(x=>({
        q:x.q||x.question, o:x.o||x.options,
        a:(x.a!==undefined?x.a:x.correct), img:normImg(x.img||x.image||'')
      })).filter(v=>v.q && Array.isArray(v.o) && v.o.length>=2 && typeof v.a==='number');
      return { title:d.title||d.t||'Quiz', questions };
    }
  }catch(e){ console.error('Chargement questions :', e.message); }
  return { title:'Quiz', questions:[] };
}
(function(){ const r=loadFromDisk(); QUIZ_TITLE=r.title; QUESTIONS=r.questions; })();

// ---------- IP locale ----------
function lanIP(){
  const nets=os.networkInterfaces();
  for(const name of Object.keys(nets)) for(const net of nets[name])
    if(net.family==='IPv4' && !net.internal) return net.address;
  return 'localhost';
}
const IP=lanIP();

// ---------- Routes : modèle + import ----------
app.get('/', (req,res)=>res.sendFile(path.join(__dirname,'public','host.html')));
app.get('/play', (req,res)=>res.sendFile(path.join(__dirname,'public','play.html')));

app.get('/template.xlsx', (req,res)=>{
  const aoa=[
    ['Question','Option A','Option B','Option C','Option D','Bonne réponse','Image (facultatif)'],
    ['Quel est le plus grand commandement ?','Aimer Dieu et son prochain','Jeûner','Donner la dîme','Prier le matin','A',''],
    ['Qui a été restauré après avoir renié Jésus ?','Judas','Pierre','Thomas','Jean','B',''],
    ['(Exemple avec image) Quel est cet animal ?','Lion','Agneau','Colombe','Aigle','C','colombe.jpg']
  ];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:44},{wch:22},{wch:16},{wch:16},{wch:16},{wch:16},{wch:20}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Questions');
  const buf=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
  res.setHeader('Content-Disposition','attachment; filename="modele_questions.xlsx"');
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

app.post('/import', express.raw({type:'*/*', limit:'15mb'}), (req,res)=>{
  try{
    const name=String(req.query.name||'').toLowerCase();
    const buf=req.body;
    let wb;
    if(name.endsWith('.csv')) wb=XLSX.read(buf.toString('utf8'),{type:'string'});
    else wb=XLSX.read(buf,{type:'buffer'});
    const qs=workbookToQuiz(wb);
    if(!qs.length) return res.status(400).json({error:'Aucune question détectée. Utilise le modèle (colonnes Question, Option A-D, Bonne réponse).'});
    QUESTIONS=qs.slice(0,50);
    if(req.query.title) QUIZ_TITLE=String(req.query.title);
    res.json({ count:QUESTIONS.length, title:QUIZ_TITLE });
  }catch(e){ res.status(400).json({error:'Fichier illisible : '+e.message}); }
});

// ---------- Jeu (une partie a la fois) ----------
let game=null;
function newPin(){ return String(Math.floor(1000+Math.random()*9000)); }

io.on('connection', (socket)=>{
  socket.on('host:create', async (payload)=>{
    const pin=newPin();
    // Adresse pour les joueurs : en LOCAL on utilise l'IP du WiFi ;
    // en LIGNE (Render, etc.) on utilise l'adresse publique du navigateur.
    const o=(payload&&payload.origin)?String(payload.origin):'';
    let base;
    if(!o || /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i.test(o)) base='http://'+IP+':'+PORT;
    else base=o.replace(/\/+$/,'');
    const joinUrl=base+'/play?pin='+pin;
    let qr=''; try{ qr=await QRCode.toDataURL(joinUrl,{margin:1,width:320}); }catch(_){}
    game={ pin, hostId:socket.id, players:{}, state:'lobby', qIndex:-1, qStart:0, timer:null };
    socket.join('host');
    socket.emit('host:created',{ pin, joinUrl, qr, total:QUESTIONS.length, title:QUIZ_TITLE, base });
  });

  socket.on('player:join', ({pin,name})=>{
    if(!game || game.state==='ended'){ socket.emit('player:error','Aucune partie en cours pour le moment.'); return; }
    if(String(pin)!==String(game.pin)){ socket.emit('player:error','Code PIN invalide.'); return; }
    name=(name||'').trim().slice(0,24)||'Joueur';
    game.players[socket.id]={ name, score:0, answered:false };
    socket.join('players');
    socket.emit('player:joined',{ name, title:QUIZ_TITLE });
    sendLobby();
  });

  socket.on('host:start', ()=>{ if(game && socket.id===game.hostId) nextQuestion(); });
  socket.on('host:next',  ()=>{ if(game && socket.id===game.hostId) nextQuestion(); });

  socket.on('player:answer', ({idx})=>{
    if(!game || game.state!=='question') return;
    const p=game.players[socket.id]; if(!p||p.answered) return;
    p.answered=true;
    const q=QUESTIONS[game.qIndex];
    const elapsed=(Date.now()-game.qStart)/1000;
    const correct=(idx===q.a);
    let gain=0;
    if(correct){ gain=500+Math.round(500*Math.max(0,(LIMIT-elapsed)/LIMIT)); p.score+=gain; }
    socket.emit('player:result',{ correct, correctIdx:q.a, gain, score:p.score });
    const total=Object.keys(game.players).length;
    const answered=Object.values(game.players).filter(x=>x.answered).length;
    io.to('host').emit('host:progress',{ answered, total });
    if(answered>=total && total>0) endQuestion();
  });

  socket.on('disconnect', ()=>{ if(game && game.players[socket.id]){ delete game.players[socket.id]; sendLobby(); } });
});

function sendLobby(){
  if(!game) return;
  const players=Object.values(game.players).map(p=>p.name);
  io.to('host').emit('host:lobby',{ players, count:players.length });
}
function nextQuestion(){
  if(!game) return;
  clearTimeout(game.timer);
  game.qIndex++;
  if(game.qIndex>=QUESTIONS.length){ endGame(); return; }
  for(const id in game.players) game.players[id].answered=false;
  game.state='question'; game.qStart=Date.now();
  const q=QUESTIONS[game.qIndex];
  io.to('host').emit('host:question',{ index:game.qIndex, total:QUESTIONS.length, q:q.q, o:q.o, img:q.img||'', limit:LIMIT });
  io.to('players').emit('player:question',{ index:game.qIndex, total:QUESTIONS.length, count:q.o.length, limit:LIMIT });
  game.timer=setTimeout(endQuestion, LIMIT*1000);
}
function endQuestion(){
  if(!game || game.state!=='question') return;
  clearTimeout(game.timer);
  game.state='reveal';
  const q=QUESTIONS[game.qIndex];
  const board=leaderboard();
  const last=game.qIndex>=QUESTIONS.length-1;
  io.to('host').emit('host:reveal',{ correctIdx:q.a, o:q.o, board, last });
  for(const id in game.players){
    const p=game.players[id];
    if(!p.answered){ p.answered=true; io.to(id).emit('player:result',{ correct:false, correctIdx:q.a, gain:0, score:p.score, timeout:true }); }
    io.to(id).emit('player:board',{ rank:rankOf(id), total:Object.keys(game.players).length, score:p.score });
  }
}
function leaderboard(){ return Object.values(game.players).map(p=>({name:p.name,score:p.score})).sort((a,b)=>b.score-a.score).slice(0,10); }
function rankOf(id){ const arr=Object.entries(game.players).map(([i,p])=>({i,score:p.score})).sort((a,b)=>b.score-a.score); return arr.findIndex(x=>x.i===id)+1; }
function endGame(){ if(!game) return; game.state='ended'; io.to('host').emit('host:ended',{ podium:leaderboard().slice(0,3), board:leaderboard() }); io.to('players').emit('player:ended',{}); }

server.listen(PORT, ()=>{
  console.log('\n==================================================');
  console.log('  QUIZ LOCAL DEMARRE');
  console.log('  Ecran animateur (cet ordinateur) : http://localhost:'+PORT);
  console.log('  Telephones (meme WiFi)           : http://'+IP+':'+PORT+'/play');
  console.log('  '+QUESTIONS.length+' questions chargees : '+QUIZ_TITLE);
  console.log('==================================================\n');
});
