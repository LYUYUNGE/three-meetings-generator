(function(){
  let meetingDefaults=null;
  const legacyApply=window.applyMeetingType;
  const voteLabels={yes:'赞成',no:'反对',abstain:'弃权',recuse:'回避'};
  const currentPeople=()=>$$('.participant-row').map(readRow).filter(x=>x.name);
  const currentShareholders=()=>currentPeople().filter(x=>x.identities.includes('股东'));
  const attendingShareholders=()=>currentShareholders().filter(x=>x.attending);
  const selectedVoteMap=article=>{
    const map={...(article._shareholderVotes||{})};
    article.querySelectorAll('.shareholder-vote-row').forEach(row=>{
      const selected=row.querySelector('input[type=radio]:checked');
      if(selected)map[row.dataset.name]=selected.value;
    });
    return map;
  };
  function voteSummary(article){
    const holders=attendingShareholders(),map=selectedVoteMap(article),totals={yes:0,no:0,abstain:0,recuse:0};
    holders.forEach(x=>totals[map[x.name]||'yes']+=Number(x.shares||0));
    const entitled=totals.yes+totals.no+totals.abstain;
    const pct=x=>entitled?`${(x*100/entitled).toFixed(4).replace(/0+$/,'').replace(/\.$/,'')}%`:'0%';
    const box=article.querySelector('.vote-summary');
    if(box)box.textContent=`有表决权股份总数：${entitled.toLocaleString('zh-CN')}股；赞成${totals.yes.toLocaleString('zh-CN')}股（${pct(totals.yes)}），反对${totals.no.toLocaleString('zh-CN')}股（${pct(totals.no)}），弃权${totals.abstain.toLocaleString('zh-CN')}股（${pct(totals.abstain)}），回避${totals.recuse.toLocaleString('zh-CN')}股。`;
  }
  function renderProposalVotes(article){
    let box=article.querySelector('.shareholder-votes');
    if(!box){box=document.createElement('div');box.className='shareholder-votes hidden';box.innerHTML='<h4>表决情况</h4><div class="shareholder-vote-list"></div><div class="vote-summary"></div>';article.querySelector('.vote-grid').after(box)}
    const shareholderMeeting=$('#meetingType').value==='shareholder';
    box.classList.toggle('hidden',!shareholderMeeting);
    article.querySelector('.vote-grid').classList.toggle('hidden',shareholderMeeting);
    article.querySelector('.p-shareholder')?.closest('label')?.classList.toggle('hidden',shareholderMeeting);
    if(!shareholderMeeting)return;
    const map=selectedVoteMap(article);article._shareholderVotes=map;
    const list=box.querySelector('.shareholder-vote-list');list.innerHTML='';
    currentShareholders().forEach((holder,index)=>{
      const result=map[holder.name]||'yes',row=document.createElement('div');row.className='shareholder-vote-row';row.dataset.name=holder.name;
      if(!holder.attending)row.classList.add('not-attending');
      row.innerHTML=`<b>${esc(holder.name)}</b><span>${Number(holder.shares||0).toLocaleString('zh-CN')}股</span>${Object.entries(voteLabels).map(([value,label])=>`<label><input type="radio" name="share_vote_${$$('.proposal').indexOf(article)}_${index}" value="${value}" ${holder.attending&&result===value?'checked':''} ${holder.attending?'':'disabled'}>${label}</label>`).join('')}`;
      row.onchange=()=>{article._shareholderVotes=selectedVoteMap(article);voteSummary(article)};list.append(row)
    });
    voteSummary(article)
  }
  function refreshAllVotes(){if($('#meetingType').value==='shareholder')$$('.proposal').forEach(renderProposalVotes)}
  function renderConvenerShareholders(){
    const list=$('#convenerShareholderList'),selected=new Set([...list.querySelectorAll('input:checked')].map(x=>x.value));list.innerHTML='';
    currentShareholders().forEach(holder=>{const label=document.createElement('label');label.innerHTML=`<input type="checkbox" name="convener_shareholder" value="${esc(holder.name)}" ${selected.has(holder.name)?'checked':''}>${esc(holder.name)}`;list.append(label)});
  }
  function syncConvener(){
    const type=$('[name=convener_type]:checked')?.value||'董事会',shareMode=type==='股东';$('#convenerShareholders').classList.toggle('hidden',!shareMode);
    if(shareMode)renderConvenerShareholders();
    const names=[...document.querySelectorAll('[name=convener_shareholder]:checked')].map(x=>x.value);
    $('[name=convener]').value=shareMode?(names.length?names.join('、'):'股东'):type;
  }
  function setCategoryChecks(type){
    const values={include_shareholder:type==='shareholder',include_board:type!=='supervisory',include_supervisory:type!=='board',include_exec:type!=='supervisory'};
    Object.entries(values).forEach(([name,value])=>{const e=$(`[name=${name}]`);if(e)e.checked=value});
    const visible=type==='supervisory'?new Set(['include_supervisory','include_other']):new Set(['include_shareholder','include_board','include_supervisory','include_exec','include_other']);
    ['include_shareholder','include_board','include_supervisory','include_exec','include_other'].forEach(name=>{const e=$(`[name=${name}]`);if(e)e.closest('label').style.display=visible.has(name)?'':'none'});
  }
  function applyType(type,resetRoster=true){
    legacyApply?.(type,resetRoster);
    setCategoryChecks(type);$('#memberCounts').classList.toggle('hidden',type==='shareholder');$('#convenerChoices').classList.toggle('hidden',type!=='shareholder');
    const convener=type==='supervisory'?'监事会':'董事会';const radio=$(`[name=convener_type][value="${convener}"]`);if(radio)radio.checked=true;
    if(type==='shareholder'){
      $$('[name=meeting_method]').forEach(e=>e.checked=true);$('[name=host]').value='王川';$('[name=host_role]').value='董事长';
      if(meetingDefaults&&resetRoster)loadRoster(meetingDefaults.shareholder.personnel)
    }
    syncConvener();renderConvenerShareholders();refreshAllVotes()
  }
  const baseProposal=proposal;
  proposal=function(v={}){baseProposal(v);const article=$$('.proposal').at(-1);article._shareholderVotes=Object.fromEntries((v.shareholder_votes||[]).map(x=>[x.name,x.result]));renderProposalVotes(article)};
  const baseCollect=collect;
  collect=function(){
    syncConvener();const data=baseCollect();
    if(data.meeting_type==='shareholder')data.proposals.forEach((proposal,index)=>{const article=$$('.proposal')[index],map=selectedVoteMap(article);proposal.shareholder_votes=attendingShareholders().map(x=>({name:x.name,shares:x.shares,result:map[x.name]||'yes'}));proposal.yes='';proposal.no='';proposal.abstain='';proposal.recuse='' });
    return data
  };
  const baseFill=fill;
  fill=function(data){baseFill(data);setTimeout(()=>{if(data?.meeting_type==='shareholder'&&data.personnel_master)loadRoster(data.personnel_master);const type=data?.convener_type||'董事会',radio=$(`[name=convener_type][value="${type}"]`);if(radio)radio.checked=true;renderConvenerShareholders();(data?.convener_shareholders||[]).forEach(name=>{const e=[...document.querySelectorAll('[name=convener_shareholder]')].find(x=>x.value===name);if(e)e.checked=true});syncConvener();refreshAllVotes()},0)};
  $('#meetingType').onchange=()=>applyType($('#meetingType').value,true);
  $('#convenerChoices').onchange=syncConvener;$('#convenerShareholderList').onchange=syncConvener;
  $('#participantList').addEventListener('change',()=>{renderConvenerShareholders();refreshAllVotes()});
  $('#participantList').addEventListener('input',()=>{renderConvenerShareholders();refreshAllVotes()});
  fetch('/api/defaults').then(r=>r.json()).then(d=>{meetingDefaults=d;applyType($('#meetingType').value,false)});
})();
