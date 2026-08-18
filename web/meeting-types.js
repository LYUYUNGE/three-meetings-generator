(function(){
  let defaults;
  const $q=s=>document.querySelector(s);
  function labelText(input,text){
    const label=input&&input.closest('label');if(!label)return;
    const node=[...label.childNodes].find(n=>n.nodeType===3&&n.textContent.trim());
    if(node)node.textContent=text;
  }
  function setParticipantIdentityVisibility(type){
    document.querySelectorAll('.participant-row').forEach(row=>{
      row.querySelectorAll('input[type=checkbox]').forEach(input=>{
        if(!['董事','监事','高级管理人员'].includes(input.value))return;
        const show=type==='supervisory'?input.value==='监事':type==='board'?input.value!=='监事':true;
        input.closest('label').style.display=show?'':'none';
      });
    });
  }
  function setVotes(value){document.querySelectorAll('.p-yes').forEach(e=>e.value=value)}
  function applyMeetingType(type,resetRoster=true){
    updateType();
    const supervisory=type==='supervisory',board=type==='board';
    const includeBoard=$q('[name=include_board]'),includeExec=$q('[name=include_exec]');
    labelText(includeBoard,supervisory?'监事':'董事');
    includeExec.closest('label').style.display=supervisory?'none':'';
    if(supervisory){includeBoard.checked=true;includeExec.checked=false}
    labelText($q('[name=expected_directors]'),supervisory?'应到监事':'应到董事');
    labelText($q('[name=actual_directors]'),supervisory?'实到监事':'实到董事');
    const note=$q('#people .section-title p');if(note)note.textContent=supervisory?'默认全体监事参会；每人可以具有多个身份':board?'默认全体董事、全体高级管理人员参会；每人可以具有多个身份':note.textContent;
    if(supervisory){
      $q('[name=method]').value='现场会议';
      $q('[name=contact_name]').value='庄义峰';$q('[name=contact_phone]').value='13661241257';
      $q('[name=convener]').value='庄义峰';$q('[name=host]').value='庄义峰';$q('[name=host_role]').value='监事会主席';
      $q('[name=expected_directors]').value=3;$q('[name=actual_directors]').value=3;setVotes(3);
      if(defaults&&resetRoster){$q('[name=term_no]').value=defaults.supervisory.term_no;$q('[name=meeting_no]').value=defaults.supervisory.meeting_no;loadRoster(defaults.supervisory.personnel)}
    }else if(board){
      $q('[name=method]').value='现场结合电子通讯方式';
      $q('[name=contact_name]').value='李婧超';$q('[name=contact_phone]').value='13426243126';
      $q('[name=convener]').value='王川';$q('[name=host]').value='王川';$q('[name=host_role]').value='董事长';
      $q('[name=expected_directors]').value=7;$q('[name=actual_directors]').value=7;setVotes(7);
      if(defaults&&resetRoster){$q('[name=term_no]').value=defaults.board.term_no;$q('[name=meeting_no]').value=defaults.board.meeting_no;loadRoster(defaults.board.personnel)}
    }
    setParticipantIdentityVisibility(type);updateDate();
  }
  fetch('/api/defaults').then(r=>r.json()).then(d=>{defaults=d;setTimeout(()=>applyMeetingType($q('#meetingType').value,false),0)});
  $q('#meetingType').onchange=()=>applyMeetingType($q('#meetingType').value,true);
  new MutationObserver(()=>setParticipantIdentityVisibility($q('#meetingType').value)).observe($q('#participantList'),{childList:true});
  window.applyMeetingType=applyMeetingType;
})();
