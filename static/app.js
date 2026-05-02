const $=id=>document.getElementById(id);
const form=$("uploadForm"),thesisInput=$("thesisInput"),schoolInput=$("schoolInput"),thesisName=$("thesisName"),schoolName=$("schoolName"),submitBtn=$("submitBtn"),statusBox=$("status"),resultBox=$("result"),summaryGrid=$("summaryGrid"),typeCounts=$("typeCounts"),groupedIssues=$("groupedIssues"),downloadLinks=$("downloadLinks"),filenameText=$("filenameText"),autoRepair=$("autoRepair"),compareBox=$("compareBox"),schoolRulesBox=$("schoolRulesBox"),thesisTile=$("thesisTile"),schoolTile=$("schoolTile");
let loadingTimer=null,loadingStep=0,lastReport=null,paymentPollTimer=null;
const loadingSteps=["📄 正在解析 Word 文档结构…","🔍 正在识别正文引用与参考文献…","🔗 正在保留已有交叉引用并补全缺失链接…","🧹 正在清理正文异常空格…","📊 正在生成检测报告与修复版 Word…"];

function bindFile(input,nameEl,tile,placeholder){
  input.addEventListener("change",()=>{
    const file=input.files&&input.files.length?input.files[0]:null;
    nameEl.textContent=file?file.name:placeholder;
    tile.classList.toggle("hasFile",!!file);
  });
}
bindFile(thesisInput,thesisName,thesisTile,"上传论文 Word");
bindFile(schoolInput,schoolName,schoolTile,"上传学校格式要求");

form.addEventListener("submit",async e=>{
  e.preventDefault();
  if(!thesisInput.files.length){showStatus("请先上传论文 Word 文件。","warn");return}
  const data=new FormData();
  data.append("thesis_file",thesisInput.files[0]);
  if(schoolInput.files.length)data.append("school_requirement_file",schoolInput.files[0]);
  const enabled=autoRepair?autoRepair.checked:true;
  const params=new URLSearchParams({fix_superscript:enabled?"true":"false",fix_citation_ranges:enabled?"true":"false",fix_school_format:enabled?"true":"false"});
  submitBtn.disabled=true;
  submitBtn.textContent="检测中…";
  resultBox.classList.add("hidden");
  startLoading();
  try{
    const response=await fetch(`/api/analyze?${params.toString()}`,{method:"POST",body:data});
    const contentType=response.headers.get("content-type")||"";
    let payload=null;
    if(contentType.includes("application/json"))payload=await response.json();
    else payload={detail:await response.text()};
    if(!response.ok)throw new Error(payload.detail||payload.message||`检测失败（${response.status}）`);
    stopLoading("✅ 检测完成，报告与修复版已生成。","done");
    lastReport=payload;
    renderReport(payload);
  }catch(error){
    stopLoading(error.message||"检测失败，请稍后重试。","warn");
  }finally{
    submitBtn.disabled=false;
    submitBtn.textContent="开始检测并生成修复版";
  }
});

function startLoading(){loadingStep=0;showStatus(loadingSteps[0],"loading");clearInterval(loadingTimer);loadingTimer=setInterval(()=>{loadingStep=Math.min(loadingStep+1,loadingSteps.length-1);showStatus(loadingSteps[loadingStep],"loading")},1300)}
function stopLoading(message,type="done"){clearInterval(loadingTimer);loadingTimer=null;showStatus(message,type)}
function showStatus(message,type="loading"){statusBox.innerHTML=`<div class="statusInner"><span class="statusDot ${type}"></span><span>${escapeHtml(message)}</span></div>${type==="loading"?`<div class="progressTrack"><span style="width:${Math.min(92,(loadingStep+1)*20)}%"></span></div>`:""}`;statusBox.classList.remove("hidden")}
function renderReport(report){resultBox.classList.remove("hidden");filenameText.textContent=report.filename||"";renderSummary(report.summary||{});renderCompare(report.before_summary,report.summary||{});renderSchoolRules(report.school_rules);renderTypeCounts(report.issue_type_counts||{});renderGroupedIssues(report.issues_by_group||{},report.group_labels||{});renderDownloads(report);resultBox.scrollIntoView({behavior:"smooth",block:"start"})}
function renderSummary(summary){const items=[["发现问题",summary.total_issues],["需确认",summary.manual_confirm_count],["格式提醒",summary.reminder_count],["学校规则",summary.school_format_count],["引用标记",summary.citation_marker_count],["连续引用",summary.citation_sequence_count],["参考文献",summary.reference_count],["已引编号",summary.cited_reference_number_count]];summaryGrid.innerHTML=items.map(([label,value])=>`<div class="metric"><div class="label">${label}</div><div class="value">${value??0}</div></div>`).join("")}
function renderCompare(before,after){if(!before){compareBox.classList.add("hidden");return}compareBox.classList.remove("hidden");compareBox.innerHTML=`<div class="panelTitle"><span>修复前后对比</span></div><div class="summaryGrid compact"><div class="metric"><div class="label">修复前问题</div><div class="value">${before.total_issues??0}</div></div><div class="metric"><div class="label">修复后问题</div><div class="value">${after.total_issues??0}</div></div><div class="metric"><div class="label">修复前可自动处理</div><div class="value">${before.auto_fixable_count??0}</div></div><div class="metric"><div class="label">修复后可自动处理</div><div class="value">${after.auto_fixable_count??0}</div></div></div>`}
function renderSchoolRules(schoolRules){if(!schoolRules||!schoolRules.rule_count){schoolRulesBox.classList.add("hidden");return}const rules=schoolRules.rules||{};const names={body:"正文",abstract_title:"中文摘要标题",abstract:"中文摘要内容",english_abstract_title:"英文摘要标题",english_abstract:"英文摘要内容",heading1:"章标题",heading2:"一级节标题",heading3:"二级节标题",figure_caption:"图题",table_caption:"表题",reference_title:"参考文献标题",reference:"参考文献正文",keywords:"中文关键词",english_keywords:"英文关键词"};const tags=Object.entries(rules).map(([category,rule])=>{const parts=[names[category]||category];if(rule.font_east_asia)parts.push(`中文 ${rule.font_east_asia}`);if(rule.font_latin)parts.push(`英文/数字 ${rule.font_latin}`);if(rule.size_name)parts.push(rule.size_name);else if(rule.size_pt)parts.push(`${rule.size_pt} 磅`);if(rule.line_spacing_name)parts.push(rule.line_spacing_name);else if(rule.line_spacing_pt)parts.push(`固定值 ${rule.line_spacing_pt} 磅`);return `<span class="ruleTag">${escapeHtml(parts.join(" / "))}</span>`}).join("");schoolRulesBox.classList.remove("hidden");schoolRulesBox.innerHTML=`<div class="panelTitle"><span>已识别学校格式要求</span><small>${schoolRules.rule_count} 类规则</small></div><div class="ruleTags">${tags}</div>`}
function renderTypeCounts(counts){const entries=Object.entries(counts);if(!entries.length){typeCounts.innerHTML=`<div class="listItem"><span>暂无问题</span><strong>0</strong></div>`;return}typeCounts.innerHTML=entries.map(([key,value])=>`<div class="listItem"><span>${escapeHtml(key)}</span><strong>${value}</strong></div>`).join("")}
function renderGroupedIssues(groups,labels){const order=["fixed","manual","reminder","school","auto"];const html=order.map(key=>{const items=groups[key]||[];if(!items.length)return"";const title=labels[key]||key;return `<div class="groupBlock"><h4 class="groupTitle">${escapeHtml(title)}（${items.length}）</h4>${items.slice(0,30).map(renderIssue).join("")}</div>`}).join("");groupedIssues.innerHTML=html||`<div class="issue empty"><strong>没有发现需要展示的问题。</strong><p>当前文件未发现明显格式问题。</p></div>`}
function renderIssue(issue){return `<div class="issue"><div class="issueTop"><strong>${escapeHtml(issue.problem)}</strong><span class="badge">${escapeHtml(issue.label||issue.type)}</span></div><p>${escapeHtml(issue.suggestion||"")}</p>${issue.text?`<p class="raw">${escapeHtml(issue.text)}</p>`:""}</div>`}
function renderDownloads(report){
  const links=[];
  const paywall=report.paywall||{};
  const paid=!paywall.enabled||paywall.paid;
  if(report.fixed&&report.fixed.download_url){
    if(paid)links.push(`<a class="primaryLink" href="${report.fixed.download_url}">下载修复版 Word</a>`);
    else links.push(`<button type="button" class="payButton" onclick="showPayPanel()">支付 ¥${escapeHtml(paywall.amount_yuan||'19.90')} 解锁修复版 Word</button>`);
  }
  if(report.report_files&&report.report_files.txt_download_url)links.push(`<a href="${report.report_files.txt_download_url}">TXT 报告</a>`);
  if(report.report_files&&report.report_files.json_download_url)links.push(`<a href="${report.report_files.json_download_url}">JSON 报告</a>`);
  downloadLinks.innerHTML=links.join("")+(!paid?renderPayPanel(report):"");
}
function renderPayPanel(report){return `<div id="payPanel" class="payPanel hidden"><div class="payTitle">解锁修复版 Word</div><div class="paySub">支付成功后将自动显示下载按钮。</div><div class="payActions"><button type="button" onclick="createPayment('wechat')">微信支付</button><button type="button" onclick="createPayment('alipay')">支付宝支付</button></div><div id="payContent" class="payContent"></div></div>`}
function showPayPanel(){const panel=$("payPanel");if(panel)panel.classList.remove("hidden")}
async function createPayment(provider){
  if(!lastReport||!lastReport.job_id){alert("未找到当前检测任务，请重新检测。");return}
  showPayPanel();
  const content=$("payContent");
  content.innerHTML="正在创建支付订单…";
  const data=new FormData();
  data.append("job_id",lastReport.job_id);
  data.append("provider",provider);
  try{
    const response=await fetch("/api/payment/create",{method:"POST",body:data});
    const payload=await response.json();
    if(!response.ok)throw new Error(payload.detail||"创建支付订单失败");
    if(payload.paid){await refreshPaymentStatus(true);return}
    if(provider==="alipay"){
      content.innerHTML=`<p>正在跳转支付宝收银台…</p><p class="payOrder">订单号：${escapeHtml(payload.out_trade_no||"")}</p>`;
      window.open(payload.pay_url,"_blank");
    }else{
      content.innerHTML=`<p>请使用微信扫码支付 ¥${escapeHtml(payload.amount_yuan||"")}</p><img class="payQr" src="${payload.qr_data_uri}" alt="微信支付二维码"><p class="payOrder">订单号：${escapeHtml(payload.out_trade_no||"")}</p>`;
    }
    startPaymentPolling();
  }catch(error){content.innerHTML=`<p class="payError">${escapeHtml(error.message||"创建支付订单失败")}</p>`}
}
function startPaymentPolling(){clearInterval(paymentPollTimer);paymentPollTimer=setInterval(()=>refreshPaymentStatus(false),2500)}
async function refreshPaymentStatus(force){
  if(!lastReport||!lastReport.job_id)return;
  const response=await fetch(`/api/payment/status?job_id=${encodeURIComponent(lastReport.job_id)}`);
  const payload=await response.json();
  if(payload.paid){
    clearInterval(paymentPollTimer);
    lastReport.paywall=lastReport.paywall||{};
    lastReport.paywall.paid=true;
    const content=$("payContent");
    if(content)content.innerHTML="✅ 支付成功，已解锁下载。";
    renderDownloads(lastReport);
  }else if(force){
    const content=$("payContent");
    if(content)content.innerHTML="尚未收到支付结果，请稍后再试。";
  }
}
function escapeHtml(value){return String(value).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}
window.showPayPanel=showPayPanel;window.createPayment=createPayment;
