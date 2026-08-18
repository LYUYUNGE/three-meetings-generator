(function(){
  const baseProposal=proposal;
  proposal=function(v={}){baseProposal(v);const article=$$('.proposal').at(-1),label=document.createElement('label');label.className='result-note-label';label.innerHTML=`审议结果说明（选填）<textarea class="p-result-note" placeholder="如需对本议案的审议结果作进一步说明，请在此填写；将分别写入会议记录和会议决议。">${esc(v.result_note||v.record_after_vote||v.resolution_after_vote||'')}</textarea>`;article.querySelector('.vote-grid').after(label)};
  const baseCollect=collect;
  collect=function(){const data=baseCollect();data.proposals.forEach((p,i)=>p.result_note=$$('.proposal')[i].querySelector('.p-result-note')?.value.trim()||'');return data};
})();
