import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Headphones, ArrowLeft, Send, Plus, MessageCircle, CheckCircle } from 'lucide-react';
import { supportAPI } from '../api/client';
import useAuthStore from '../stores/authStore';
const ST={open:'bg-red-500/20 text-red-400',in_progress:'bg-yellow-500/20 text-yellow-400',closed:'bg-green-500/20 text-green-400'};
const STR={open:'Открыт',in_progress:'В работе',closed:'Решён'};
export default function SupportRequest(){
  const[tickets,setTickets]=useState([]);const[sel,setSel]=useState(null);const[reply,setReply]=useState('');
  const[showNew,setShowNew]=useState(false);const[form,setForm]=useState({subject:'',message:'',category:'general'});
  const navigate=useNavigate();const user=useAuthStore(s=>s.user);
  useEffect(()=>{load()},[]);
  const load=()=>supportAPI.tickets().then(({data})=>setTickets(data)).catch(()=>{});
  const create=async()=>{if(!form.subject.trim()||!form.message.trim())return;await supportAPI.createTicket(form);setShowNew(false);setSel(null);setForm({subject:'',message:'',category:'general'});load()};
  const sendReply=async()=>{if(!reply.trim()||!sel)return;await supportAPI.reply(sel.id,reply);setReply('');load();setSel(p=>({...p,replies:[...(p.replies||[]),{user_name:user.full_name,message:reply,user_role:user.role}]}))};
  return(<div className="min-h-screen bg-gray-900 text-white flex">
    <div className="w-80 border-r border-gray-700 flex flex-col">
      <div className="px-4 py-3 border-b border-gray-700"><div className="flex items-center justify-between"><h2 className="font-bold flex items-center gap-2"><Headphones className="w-5 h-5 text-purple-400"/>Техподдержка</h2><button onClick={()=>navigate(-1)} className="text-gray-400 hover:text-white text-sm">Назад</button></div></div>
      <div className="p-3 border-b border-gray-700"><button onClick={()=>{setShowNew(true);setSel(null)}} className="w-full py-2 bg-purple-600 hover:bg-purple-700 rounded-lg text-sm flex items-center justify-center gap-1"><Plus className="w-4 h-4"/>Новое обращение</button></div>
      <div className="flex-1 overflow-y-auto">{tickets.length===0&&<div className="text-center py-10 text-gray-500 text-sm">Нет обращений</div>}
        {tickets.map(t=>(<button key={t.id} onClick={()=>setSel(t)} className={`w-full text-left px-4 py-3 border-b border-gray-800 hover:bg-gray-800 ${sel?.id===t.id?'bg-gray-800':''}`}>
          <div className="flex items-center justify-between mb-1"><span className="font-medium text-sm truncate flex-1">{t.subject}</span><span className={`text-[10px] px-1.5 py-0.5 rounded ml-2 ${ST[t.status]}`}>{STR[t.status]||t.status}</span></div>
          <span className="text-xs text-gray-500">{new Date(t.created_at).toLocaleDateString('ru')}</span></button>))}</div></div>
    <div className="flex-1 flex flex-col">{sel?(<>
      <div className="px-5 py-3 border-b border-gray-700"><h3 className="font-semibold">{sel.subject}</h3><span className={`text-xs px-2 py-0.5 rounded ${ST[sel.status]}`}>{STR[sel.status]}</span></div>
      <div className="flex-1 overflow-y-auto p-4 space-y-3">
        <div className="flex justify-end"><div className="max-w-[70%] px-4 py-3 rounded-2xl bg-blue-600 rounded-br-md"><div className="text-sm whitespace-pre-wrap">{sel.message}</div><div className="text-[10px] text-blue-300 mt-1">{new Date(sel.created_at).toLocaleString('ru')}</div></div></div>
        {sel.replies?.map((r,i)=>(<div key={i} className={`flex ${r.user_role==='support'||r.user_role==='admin'?'justify-start':'justify-end'}`}>
          <div className={`max-w-[70%] px-4 py-2.5 rounded-2xl text-sm ${r.user_role==='support'||r.user_role==='admin'?'bg-purple-600 rounded-bl-md':'bg-blue-600 rounded-br-md'}`}>
            {(r.user_role==='support'||r.user_role==='admin')&&<div className="text-xs opacity-70 mb-1 flex items-center gap-1"><Headphones className="w-3 h-3"/>{r.user_name}</div>}
            <div className="whitespace-pre-wrap">{r.message}</div></div></div>))}</div>
      {sel.status!=='closed'&&<div className="p-3 border-t border-gray-700 flex gap-2"><input value={reply} onChange={e=>setReply(e.target.value)} onKeyDown={e=>e.key==='Enter'&&sendReply()} className="flex-1 bg-gray-800 border border-gray-700 rounded-full px-4 py-2.5 text-sm focus:outline-none focus:border-purple-500" placeholder="Написать..."/><button onClick={sendReply} disabled={!reply.trim()} className="p-2.5 bg-purple-600 rounded-full hover:bg-purple-700 disabled:opacity-50"><Send className="w-4 h-4"/></button></div>}
    </>):showNew?(<div className="flex-1 p-6"><h2 className="text-xl font-bold mb-4">Новое обращение</h2>
      <div className="max-w-lg space-y-4">
        <div><label className="text-sm text-gray-400 block mb-1">Тема</label><input value={form.subject} onChange={e=>setForm({...form,subject:e.target.value})} className="w-full px-4 py-2 bg-gray-800 border border-gray-700 rounded-lg text-sm focus:border-purple-500 focus:outline-none" placeholder="Кратко опишите проблему"/></div>
        <div><label className="text-sm text-gray-400 block mb-1">Категория</label><select value={form.category} onChange={e=>setForm({...form,category:e.target.value})} className="w-full px-4 py-2 bg-gray-800 border border-gray-700 rounded-lg text-sm"><option value="general">Общий вопрос</option><option value="bug">Ошибка</option><option value="feature">Предложение</option><option value="account">Аккаунт</option></select></div>
        <div><label className="text-sm text-gray-400 block mb-1">Описание</label><textarea value={form.message} onChange={e=>setForm({...form,message:e.target.value})} className="w-full h-32 px-4 py-2 bg-gray-800 border border-gray-700 rounded-lg text-sm resize-none focus:border-purple-500 focus:outline-none" placeholder="Подробно опишите вашу проблему..."/></div>
        <div className="flex gap-2"><button onClick={()=>setShowNew(false)} className="px-4 py-2 bg-gray-700 rounded-lg text-sm">Отмена</button><button onClick={create} disabled={!form.subject.trim()||!form.message.trim()} className="px-4 py-2 bg-purple-600 hover:bg-purple-700 rounded-lg text-sm disabled:opacity-50">Отправить</button></div></div>
    </div>):(<div className="flex-1 flex items-center justify-center text-gray-500"><div className="text-center"><Headphones className="w-16 h-16 mx-auto mb-4 opacity-20"/><p className="text-lg mb-2">Техподдержка AI Interview Platform</p><p className="text-sm">Создайте обращение или выберите из списка</p><button onClick={()=>setShowNew(true)} className="mt-4 px-6 py-2 bg-purple-600 hover:bg-purple-700 rounded-lg text-sm text-white"><Plus className="w-4 h-4 inline mr-1"/>Новое обращение</button></div></div>)}</div>
  </div>);
}
