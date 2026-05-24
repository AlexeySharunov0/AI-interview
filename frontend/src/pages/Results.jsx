import React, { useState, useEffect } from 'react';
import { useParams, Link } from 'react-router-dom';
import { Trophy, ArrowLeft, CheckCircle, XCircle, AlertCircle } from 'lucide-react';
import { interviewAPI } from '../api/client';
export default function Results() {
  const {interviewId}=useParams();const[iv,setIv]=useState(null);const[report,setReport]=useState(null);
  useEffect(()=>{
    interviewAPI.get(interviewId).then(({data})=>setIv(data)).catch(()=>{});
    try{interviewAPI.report(interviewId).then(({data})=>setReport(data)).catch(()=>setReport(null))}catch(e){setReport(null)}
  },[interviewId]);
  if(!iv)return<div className="min-h-screen bg-gray-900 flex items-center justify-center text-white">Загрузка результатов...</div>;
  const ts=iv.total_score||0;
  const sc=ts>=70?'text-green-400':ts>=40?'text-yellow-400':'text-red-400';
  const sl=ts>=70?'Отлично!':ts>=40?'Неплохо':'Нужна подготовка';
  const getScoreColor=(s)=>(s||0)>=70?'text-green-400':(s||0)>=40?'text-yellow-400':'text-red-400';
  const getScoreBg=(s)=>(s||0)>=70?'bg-green-500/10 border-green-500/30':(s||0)>=40?'bg-yellow-500/10 border-yellow-500/30':'bg-red-500/10 border-red-500/30';
  return(<div className="min-h-screen bg-gray-900 text-white">
    <div className="px-6 py-3 bg-gray-800 border-b border-gray-700 flex items-center gap-3">
      <Link to="/interviews" className="flex items-center gap-2 text-gray-400 hover:text-white"><ArrowLeft className="w-4 h-4"/>Назад</Link>
      <h1 className="text-lg font-bold flex-1">Результаты собеседования</h1>
    </div>
    <div className="p-6">
      <div className="max-w-3xl mx-auto">
        <div className="text-center mb-8"><Trophy className={`w-14 h-14 mx-auto mb-3 ${sc}`}/><h2 className="text-2xl font-bold mb-1">Интервью завершено</h2><p className="text-gray-400 text-sm">Уровень: {iv.level?.toUpperCase()}</p></div>
        <div className="bg-gray-800 rounded-xl p-6 text-center mb-6">
          <div className={`text-5xl font-extrabold mb-1 ${sc}`}>{ts.toFixed(0)}%</div>
          <div className="text-lg text-gray-300">{sl}</div>
          <div className="text-gray-500 mt-1 text-sm">Выполнено задач: {iv.completed_tasks}/{iv.total_tasks}</div>
        </div>
        {iv.ai_recommendation&&<div className="bg-gray-800 rounded-xl p-4 mb-6 flex items-center gap-3">
          {iv.ai_recommendation==='hire'?<CheckCircle className="w-6 h-6 text-green-400"/>:iv.ai_recommendation==='maybe'?<AlertCircle className="w-6 h-6 text-yellow-400"/>:<XCircle className="w-6 h-6 text-red-400"/>}
          <div><span className="font-semibold text-sm">Рекомендация: </span>
            <span className={iv.ai_recommendation==='hire'?'text-green-400':iv.ai_recommendation==='maybe'?'text-yellow-400':'text-red-400'}>
              {iv.ai_recommendation==='hire'?'Рекомендован':iv.ai_recommendation==='maybe'?'На рассмотрении':'Не рекомендован'}</span></div></div>}
        {report?.tasks&&<div className="mb-6">
          <h3 className="text-lg font-bold mb-3">Детализация по задачам</h3>
          <div className="space-y-3">
            {report.tasks.map((t,i)=>(
              <div key={t.id} className={`rounded-xl p-4 border ${getScoreBg(t.score)}`}>
                <div className="flex justify-between items-start mb-2">
                  <div className="flex-1">
                    <span className="text-xs text-gray-500 mr-2">#{i+1}</span>
                    <span className="font-semibold text-sm">{t.title}</span>
                    {t.domain&&<span className="ml-2 text-xs px-1.5 py-0.5 bg-gray-700 rounded">{t.domain}</span>}
                    {t.level&&<span className="ml-1 text-xs px-1.5 py-0.5 bg-gray-700 rounded">{t.level}</span>}
                  </div>
                  <span className={`text-lg font-bold ${getScoreColor(t.score)}`}>{(t.score||0).toFixed(0)}%</span>
                </div>
                {t.description&&<p className="text-xs text-gray-400 mb-2">{t.description}</p>}
                <div className="flex items-center gap-4 text-xs text-gray-500">
                  <span>{t.feedback}</span>
                  {(t.score||0)>=70&&<span className="text-green-400 flex items-center gap-1"><CheckCircle className="w-3 h-3"/>Отлично</span>}
                  {(t.score||0)>0&&(t.score||0)<70&&<span className="text-yellow-400 flex items-center gap-1"><AlertCircle className="w-3 h-3"/>Частично</span>}
                  {(t.score||0)===0&&<span className="text-red-400 flex items-center gap-1"><XCircle className="w-3 h-3"/>Не пройдено</span>}
                </div>
              </div>
            ))}
          </div>
        </div>}
        {iv.ai_summary&&<div className="bg-gray-800 rounded-xl p-5 mb-6"><h3 className="font-semibold mb-2 text-sm">Общая оценка AI</h3><p className="text-gray-300 text-sm">{iv.ai_summary}</p></div>}
        <div className="text-center pb-6"><Link to="/interviews" className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 rounded-lg inline-block text-sm">Ещё одно интервью</Link></div>
      </div>
    </div>
  </div>);
}
