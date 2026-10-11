const VERSION = "FRIENDZONE-EVAL-2026-10-v1";
const PROTOTYPE_URL = "https://yueran-ml.github.io/7285/";
const TOTAL_STEPS = 6;
const STORAGE_KEY = "friendzone-evaluation-draft-v1";

const randomCode = (prefix) => {
  const bytes = new Uint8Array(4);
  crypto.getRandomValues(bytes);
  return `${prefix}-${[...bytes].map((value) => value.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
};

const freshAnswers = () => ({
  response_code: randomCode("FZ"),
  receipt_code: randomCode("RCPT"),
  version: VERSION,
  language: "en",
  started_at: new Date().toISOString(),
});

let answers = loadDraft();
let step = 0;
let submitting = false;
let submitted = false;

const copy = {
  en: {
    languageLabel: "Language / 语言",
    eyebrow: "FRIENDZONE EVALUATION",
    titles: [
      "Before you begin",
      "Task 1 · Find your circle",
      "Task 2 · Start a conversation",
      "Task 3 · Meet in person",
      "SUS and overall feedback",
      "Review and submit",
    ],
    subtitles: [
      "About 15–20 minutes. Work at your own pace.",
      "Explore recommendations and choose a group that feels right.",
      "Try the conversation support inside a circle.",
      "Test the simulated phone tap, timer and seven-day board.",
      "Rate the whole prototype after completing the tasks.",
      "Nothing is sent until you press Submit.",
    ],
    step: "Step",
    of: "of",
    back: "Back",
    next: "Continue",
    review: "Review answers",
    submit: "Submit response",
    submitting: "Submitting…",
    draftSaved: "Draft saved on this device",
    footer: "Your answers are submitted only when you press Submit.",
    requiredError: "Please confirm both eligibility and voluntary consent before continuing.",
    submitError: "The response could not be submitted. Your draft is still on this device. Check the study configuration or try again.",
    missingSus: "Some SUS items are unanswered. You may go back to complete them, or submit without a SUS score.",
  },
  zh: {
    languageLabel: "语言 / Language",
    eyebrow: "FRIENDZONE 评估研究",
    titles: [
      "开始前，请先阅读",
      "任务 1 · 找到适合你的小组",
      "任务 2 · 开始一段对话",
      "任务 3 · 参加线下聚会",
      "SUS 与整体反馈",
      "检查并提交",
    ],
    subtitles: [
      "约 15–20 分钟，请按自己的节奏完成。",
      "查看推荐，并选择一个适合你的小组。",
      "体验小组中的对话辅助功能。",
      "测试模拟手机接触、计时和七天排行榜。",
      "完成任务后，对整个原型进行评价。",
      "只有点击“提交”后，答案才会发送。",
    ],
    step: "第",
    of: "步，共",
    back: "上一步",
    next: "继续",
    review: "检查答案",
    submit: "提交回答",
    submitting: "正在提交…",
    draftSaved: "草稿已保存在本设备",
    footer: "只有点击“提交”后，答案才会发送。",
    requiredError: "请先确认参与条件和自愿同意。",
    submitError: "提交失败。草稿仍保存在本设备，请检查研究配置或稍后重试。",
    missingSus: "部分 SUS 题目尚未回答。你可以返回补充，也可以提交但不生成 SUS 分数。",
  },
};

const susItems = {
  en: [
    "I think that I would like to use FriendZone frequently.",
    "I found FriendZone unnecessarily complex.",
    "I thought FriendZone was easy to use.",
    "I think that I would need the support of a technical person to use FriendZone.",
    "I found the various functions in FriendZone were well integrated.",
    "I thought there was too much inconsistency in FriendZone.",
    "I imagine that most people would learn to use FriendZone very quickly.",
    "I found FriendZone very cumbersome to use.",
    "I felt very confident using FriendZone.",
    "I needed to learn a lot of things before I could get going with FriendZone.",
  ],
  zh: [
    "我认为我会愿意经常使用 FriendZone。",
    "我觉得 FriendZone 不必要地复杂。",
    "我认为 FriendZone 很容易使用。",
    "我认为需要技术人员的帮助才能使用 FriendZone。",
    "我认为 FriendZone 的各项功能整合得很好。",
    "我认为 FriendZone 存在太多不一致之处。",
    "我认为大多数人都能很快学会使用 FriendZone。",
    "我觉得 FriendZone 使用起来很麻烦。",
    "我在使用 FriendZone 时感到很有信心。",
    "在开始使用 FriendZone 前，我需要学习很多东西。",
  ],
};

const content = document.querySelector("#content");
const errorBox = document.querySelector("#form-error");
const backButton = document.querySelector("#back");
const nextButton = document.querySelector("#next");
const languageSelect = document.querySelector("#language");

languageSelect.value = answers.language || "en";

function loadDraft() {
  try {
    const draft = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (draft && draft.version === VERSION && draft.response_code) return draft;
  } catch (_) {
    // A corrupt or blocked localStorage entry should not stop participation.
  }
  return freshAnswers();
}

function saveDraft() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(answers));
    document.querySelector("#draft-state").textContent = copy[answers.language].draftSaved;
  } catch (_) {
    document.querySelector("#draft-state").textContent = "";
  }
}

const esc = (value = "") => String(value)
  .replaceAll("&", "&amp;")
  .replaceAll("<", "&lt;")
  .replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;")
  .replaceAll("'", "&#039;");

function field(key, label, hint = "") {
  return `<label class="field">
    <span class="field-label">${esc(label)}</span>
    <textarea data-key="${esc(key)}">${esc(answers[key] || "")}</textarea>
    ${hint ? `<span class="hint">${esc(hint)}</span>` : ""}
  </label>`;
}

function selectField(key, label, options) {
  return `<label class="field">
    <span class="field-label">${esc(label)}</span>
    <select data-key="${esc(key)}">
      ${options.map(([value, text]) => `<option value="${esc(value)}" ${answers[key] === value ? "selected" : ""}>${esc(text)}</option>`).join("")}
    </select>
  </label>`;
}

function checkbox(key, label) {
  return `<label class="check">
    <input type="checkbox" data-key="${esc(key)}" ${answers[key] ? "checked" : ""} />
    <span>${esc(label)}</span>
  </label>`;
}

function taskFields(number, labels) {
  const statuses = answers.language === "zh"
    ? [["", "请选择／跳过"], ["independent", "独立完成"], ["assisted", "在帮助下完成"], ["not_completed", "尝试了但未完成"], ["stopped", "主动停止"], ["access_failed", "无法访问原型"]]
    : [["", "Please choose / skip"], ["independent", "Completed independently"], ["assisted", "Completed with help"], ["not_completed", "Tried but could not complete"], ["stopped", "Decided to stop"], ["access_failed", "Could not access the prototype"]];
  return `${selectField(`task_${number}_status`, labels.status, statuses)}
    ${field(`task_${number}_choice`, labels.choice, labels.choiceHint)}
    ${field(`task_${number}_difficulty`, labels.difficulty, labels.difficultyHint)}`;
}

function renderIntro() {
  const zh = answers.language === "zh";
  return `<section class="card">
    <h2>${zh ? "研究说明" : "Study information"}</h2>
    <p>${zh
      ? "本研究评估 FriendZone 原型是否容易使用，以及它的兴趣匹配、对话提示、模拟 NFC 接触和排行榜如何影响用户参与社交活动的感受。你将完成三个任务、SUS 问卷和几道开放题。我们测试的是界面，而不是你的能力；不需要给出正面评价。"
      : "This study evaluates whether the FriendZone prototype is easy to use and how its interest matching, conversation prompts, simulated NFC tap and leaderboard affect people’s willingness to participate. You will try three tasks, complete the SUS questionnaire and answer several open questions. We are testing the interface, not you; positive feedback is not expected."}</p>
    <p>${zh
      ? "Prototype 中的用户、小组、聚会、积分和 NFC 记录均为模拟数据。手机之间不会发生真实 NFC 传输，也不会读取位置或通讯录。"
      : "The users, circles, meetups, points and NFC records in the prototype are simulated. No real phone-to-phone NFC transfer occurs, and the study does not access your location or contacts."}</p>
    <p>${zh
      ? "参与完全自愿。你可以跳过非必答问题、暂停或停止，无需说明原因。不要在开放回答中填写姓名、学号、地址或其他身份信息。"
      : "Participation is voluntary. You may skip optional questions, pause or stop without giving a reason. Do not include names, student numbers, addresses or other identifying details in open-text answers."}</p>
    <p>${zh
      ? "提交后，回答、随机编号、问卷版本和提交时间会保存到研究数据库。匿名参与者只能新增回答，无法读取数据库。托管服务可能处理正常的技术日志。经你单独同意后，论文可引用简短且去身份化的原话。"
      : "On submission, your answers, random response code, questionnaire version and submission time are stored in the study database. Anonymous participants may add a response but cannot read the database. Hosting services may process normal technical logs. Short de-identified quotations may be used only with your separate permission."}</p>
    <p>${zh ? "你的回答编号" : "Your response code"}: <span class="code">${esc(answers.response_code)}</span></p>
  </section>
  <section class="card">
    ${checkbox("eligible", zh ? "我已满 18 岁，并且目前在布里斯班学习或生活。" : "I am 18 or older and currently study or live in Brisbane.")}
    ${checkbox("consent", zh ? "我已阅读以上说明，自愿参与并同意将反馈用于本课程研究。" : "I have read the information above and voluntarily agree to take part and share my feedback for this course research.")}
    ${selectField("quote_permission", zh ? "是否允许使用简短、去身份化的原话？" : "May short de-identified quotations from your written answers be used?", zh
      ? [["", "请选择／跳过"], ["yes", "可以"], ["no", "不可以"]]
      : [["", "Please choose / skip"], ["yes", "Yes"], ["no", "No"]])}
    ${selectField("device", zh ? "测试所使用的设备" : "Device used for the prototype", zh
      ? [["", "请选择／跳过"], ["phone", "手机"], ["tablet", "平板"], ["desktop", "笔记本／台式电脑"], ["other", "其他"]]
      : [["", "Please choose / skip"], ["phone", "Phone"], ["tablet", "Tablet"], ["desktop", "Laptop / desktop"], ["other", "Other"]])}
    ${selectField("prior_use", zh ? "你之前使用过 FriendZone Prototype 吗？" : "Have you used the FriendZone prototype before?", zh
      ? [["", "请选择／跳过"], ["no", "没有"], ["yes", "使用过"]]
      : [["", "Please choose / skip"], ["no", "No"], ["yes", "Yes"]])}
  </section>`;
}

function renderTask(number) {
  const zh = answers.language === "zh";
  const tasks = zh ? [
    {
      prompt: "想象你刚来到布里斯班。设置你的兴趣和任何你愿意分享的可选资料，然后找到一个你在现实中会考虑加入的小组。请注意该小组为什么被推荐给你。",
      status: "这个任务的完成情况是？",
      choice: "你选择了哪个小组？为什么？",
      choiceHint: "如果无法选择，请说明缺少了哪些信息。",
      difficulty: "哪里不清楚或不好操作？有人帮助你吗？",
      difficultyHint: "简单举例即可；没有困难也可以写“没有”。",
    },
    {
      prompt: "加入一个小组，并使用系统提供的 conversation prompt 发起一段对话。尝试查看另一个提示，并说说你是否愿意在真实小组中发送这样的内容。",
      status: "这个任务的完成情况是？",
      choice: "你使用了哪个提示？它是否让开场变得更容易？为什么？",
      choiceHint: "请评价提示是否自然、相关、过于私人或过于刻意。",
      difficulty: "哪里不清楚或不好操作？有人帮助你吗？",
      difficultyHint: "简单举例即可；没有困难也可以写“没有”。",
    },
    {
      prompt: "确认参加小组聚会，进入聚会模式，并通过点击或拖动模拟两部手机接触。让计时器运行，结束聚会，查看 summary 和 Board，并切换 Time together / People met。",
      status: "这个任务的完成情况是？",
      choice: "你认为手机接触、计时器和排行榜分别记录了什么？",
      choiceHint: "不确定也可以如实说明。",
      difficulty: "哪里不清楚或不好操作？有人帮助你吗？",
      difficultyHint: "请特别说明 NFC、计时开始/结束或排行榜是否清楚。",
    },
  ] : [
    {
      prompt: "Imagine you have recently arrived in Brisbane. Set up your interests and any optional information you feel comfortable sharing, then find a circle you would realistically consider joining. Notice why it was recommended.",
      status: "What happened in this task?",
      choice: "Which circle did you choose, and why?",
      choiceHint: "If you could not choose, explain what information was missing.",
      difficulty: "Was anything confusing or difficult? Did anyone help you?",
      difficultyHint: "A short example is enough; ‘nothing’ is also fine.",
    },
    {
      prompt: "Join a circle and use one of the built-in conversation prompts to start a conversation. Try another prompt and consider whether you would send something like this in a real group.",
      status: "What happened in this task?",
      choice: "Which prompt did you use? Did it make starting easier? Why or why not?",
      choiceHint: "Consider whether it felt natural, relevant, too personal or too scripted.",
      difficulty: "Was anything confusing or difficult? Did anyone help you?",
      difficultyHint: "A short example is enough; ‘nothing’ is also fine.",
    },
    {
      prompt: "Confirm that you are attending a meetup, enter meetup mode, and simulate two phones touching by clicking or dragging. Let the timer run, end the meetup, view the summary and Board, then switch between Time together and People met.",
      status: "What happened in this task?",
      choice: "What did you think the phone tap, timer and leaderboard were recording?",
      choiceHint: "It is fine to say that you were unsure.",
      difficulty: "Was anything confusing or difficult? Did anyone help you?",
      difficultyHint: "Please mention whether NFC, timing or the leaderboard was unclear.",
    },
  ];
  const task = tasks[number - 1];
  return `<section class="card">
    <div class="task-prompt"><strong>${zh ? `任务 ${number}` : `TASK ${number}`}</strong>${esc(task.prompt)}</div>
    <p class="muted">${zh
      ? number === 1 ? "请在新标签页打开原型。如果你以前使用过它，可先在 Demo controls 或 Me 页面选择 Reset prototype。三个任务应连续完成，不要在任务之间重置。" : "返回同一个原型标签页继续操作。三个任务应连续完成。"
      : number === 1 ? "Open the prototype in a new tab. If you have used it before, choose Reset prototype in the demo controls or Me screen first. Complete all three tasks as one continuous journey; do not reset between tasks." : "Return to the same prototype tab and continue. The tasks form one continuous journey."}</p>
    <a class="button button-sage prototype-link" href="${PROTOTYPE_URL}" target="_blank" rel="noopener">${zh ? "在新标签页打开 Prototype" : "Open prototype in a new tab"}</a>
    ${taskFields(number, task)}
  </section>`;
}

function renderScaleItem(text, index) {
  const key = `sus_${index + 1}`;
  const value = Number(answers[key] || 0);
  const zh = answers.language === "zh";
  return `<fieldset class="scale">
    <legend class="scale-title">${index + 1}. ${esc(text)}</legend>
    <div class="scale-options">
      ${[1, 2, 3, 4, 5].map((score) => `<span class="scale-option"><input type="radio" id="${key}_${score}" name="${key}" data-key="${key}" value="${score}" ${value === score ? "checked" : ""}><label for="${key}_${score}">${score}</label></span>`).join("")}
    </div>
    <div class="scale-anchors"><span>${zh ? "非常不同意" : "Strongly disagree"}</span><span>${zh ? "非常同意" : "Strongly agree"}</span></div>
  </fieldset>`;
}

function renderFeedback() {
  const zh = answers.language === "zh";
  const score = calculateSus();
  return `<section class="card">
    <h2>System Usability Scale (SUS)</h2>
    <p class="muted">${zh ? "请根据整个 FriendZone Prototype 作答。1 = 非常不同意，5 = 非常同意。要计算标准 SUS 分数，需要回答全部十题。" : "Answer for the FriendZone prototype as a whole. 1 = strongly disagree and 5 = strongly agree. All ten items are needed to calculate a standard SUS score."}</p>
    ${susItems[answers.language].map(renderScaleItem).join("")}
    ${score === null ? "" : `<div class="score"><span>${zh ? "当前 SUS 分数" : "Current SUS score"}</span><strong>${score}</strong><span>/ 100</span></div>`}
  </section>
  <section class="card">
    <h2>${zh ? "体验与设计反馈" : "Experience and design feedback"}</h2>
    ${field("most_helpful", zh ? "哪个功能最有帮助？为什么？" : "Which feature was most helpful, and why?")}
    ${field("priority_change", zh ? "如果只能修改一处，你最希望改什么？" : "If you could change one thing, what would it be?")}
    ${field("board_reaction", zh ? "看到七天排行榜时，你的第一反应是什么？它会激励你、给你压力，还是两者都有？" : "What was your first reaction to the seven-day leaderboard? Would it motivate you, pressure you, or both?")}
    ${field("nfc_privacy", zh ? "在真实聚会中使用手机接触并记录相处时间，会让你有什么感受或隐私顾虑？" : "How would you feel about tapping phones and recording time together at a real meetup? What privacy concerns, if any, would you have?")}
    ${field("continuity", zh ? "会后七天积分奖励会让你继续联系小组，还是会感觉被催促？真正促使你再次联系他们的会是什么？" : "Would the seven-day post-meetup bonus help you keep talking, or feel like pressure? What would genuinely make you contact the group again?")}
  </section>`;
}

function answerText() {
  const zh = answers.language === "zh";
  const labels = zh ? {
    response_code: "回答编号", device: "设备", prior_use: "是否使用过原型", quote_permission: "引用许可",
    task_1_status: "任务 1 状态", task_1_choice: "任务 1 选择", task_1_difficulty: "任务 1 困难",
    task_2_status: "任务 2 状态", task_2_choice: "任务 2 选择", task_2_difficulty: "任务 2 困难",
    task_3_status: "任务 3 状态", task_3_choice: "任务 3 理解", task_3_difficulty: "任务 3 困难",
    most_helpful: "最有帮助", priority_change: "优先修改", board_reaction: "排行榜反应", nfc_privacy: "NFC 与隐私", continuity: "会后联系",
  } : {
    response_code: "Response code", device: "Device", prior_use: "Prior prototype use", quote_permission: "Quotation permission",
    task_1_status: "Task 1 status", task_1_choice: "Task 1 choice", task_1_difficulty: "Task 1 difficulty",
    task_2_status: "Task 2 status", task_2_choice: "Task 2 response", task_2_difficulty: "Task 2 difficulty",
    task_3_status: "Task 3 status", task_3_choice: "Task 3 interpretation", task_3_difficulty: "Task 3 difficulty",
    most_helpful: "Most helpful", priority_change: "Priority change", board_reaction: "Leaderboard reaction", nfc_privacy: "NFC and privacy", continuity: "Post-meetup continuity",
  };
  const lines = [`FriendZone Evaluation · ${VERSION}`];
  Object.entries(labels).forEach(([key, label]) => lines.push(`${label}: ${answers[key] || "—"}`));
  for (let index = 1; index <= 10; index += 1) lines.push(`SUS ${index}: ${answers[`sus_${index}`] || "—"}`);
  lines.push(`SUS score: ${calculateSus() ?? "Incomplete"}`);
  return lines.join("\n");
}

function renderReview() {
  const zh = answers.language === "zh";
  const score = calculateSus();
  return submitted ? `<section class="card">
    <h2>${zh ? "回答已保存" : "Response saved"}</h2>
    <p>${zh ? "请截图或保存以下回执编号。如需更正或撤回回答，请把回执编号发给研究团队。" : "Save or screenshot the receipt below. To request a correction or withdrawal, send this receipt to the research team."}</p>
    <p class="code">${esc(answers.response_code)}<br>${esc(answers.receipt_code)}</p>
    <p class="notice">${zh ? "现在可以安全关闭此页面。" : "You may now safely close this page."}</p>
  </section>` : `<section class="card">
    <h2>${zh ? "提交前检查" : "Check before submitting"}</h2>
    <p>${zh ? "你的回答尚未发送。你可以返回修改，或下载一份个人副本。" : "Your answers have not been sent. You can go back to edit them or download a personal copy."}</p>
    ${score === null ? `<p class="notice warning">${copy[answers.language].missingSus}</p>` : `<div class="score"><span>SUS</span><strong>${score}</strong><span>/ 100</span></div>`}
    <pre class="review" id="review-output"></pre>
    <div class="row">
      <button class="button button-secondary" id="download" type="button">${zh ? "下载回答 (.txt)" : "Download answers (.txt)"}</button>
    </div>
    <p id="database-note" class="hint"></p>
  </section>`;
}

function bindInputs() {
  content.querySelectorAll("[data-key]").forEach((control) => {
    const eventName = control.matches("textarea") ? "input" : "change";
    control.addEventListener(eventName, () => {
      const key = control.dataset.key;
      if (control.type === "checkbox") answers[key] = control.checked;
      else if (control.type === "radio") answers[key] = Number(control.value);
      else answers[key] = control.value;
      saveDraft();
      if (step === 4 && control.type === "radio") render();
    });
  });
  document.querySelector("#download")?.addEventListener("click", downloadAnswers);
}

function calculateSus() {
  const values = Array.from({ length: 10 }, (_, index) => Number(answers[`sus_${index + 1}`] || 0));
  if (values.some((value) => value < 1 || value > 5)) return null;
  const contribution = values.reduce((total, value, index) => total + (index % 2 === 0 ? value - 1 : 5 - value), 0);
  return contribution * 2.5;
}

function render() {
  const lang = answers.language || "en";
  const t = copy[lang];
  document.documentElement.lang = lang === "zh" ? "zh-CN" : "en";
  document.querySelector("[data-i18n='languageLabel']").textContent = t.languageLabel;
  document.querySelector("#eyebrow").textContent = t.eyebrow;
  document.querySelector("#page-title").textContent = t.titles[step];
  document.querySelector("#page-subtitle").textContent = t.subtitles[step];
  document.querySelector("#step-label").textContent = lang === "zh" ? `${t.step} ${step + 1} ${t.of} ${TOTAL_STEPS} 步` : `${t.step} ${step + 1} ${t.of} ${TOTAL_STEPS}`;
  document.querySelector("#progress-fill").style.width = `${((step + 1) / TOTAL_STEPS) * 100}%`;
  document.querySelector(".progress-track").setAttribute("aria-valuenow", String(step + 1));
  document.querySelector("#footer-copy").textContent = t.footer;
  errorBox.textContent = "";

  if (step === 0) content.innerHTML = renderIntro();
  else if (step <= 3) content.innerHTML = renderTask(step);
  else if (step === 4) content.innerHTML = renderFeedback();
  else content.innerHTML = renderReview();

  backButton.hidden = step === 0 || submitted;
  nextButton.hidden = submitted;
  nextButton.disabled = submitting;
  nextButton.textContent = step === 5 ? (submitting ? t.submitting : t.submit) : step === 4 ? t.review : t.next;
  bindInputs();
  if (step === 5 && !submitted) {
    document.querySelector("#review-output").textContent = answerText();
    const config = window.__EVALUATION_CONFIG__ || {};
    document.querySelector("#database-note").textContent = config.supabaseUrl && config.supabaseAnonKey
      ? (lang === "zh" ? "数据库连接已配置。提交后会显示回执编号。" : "Database connection is configured. A receipt will appear after submission.")
      : (lang === "zh" ? "研究数据库尚未配置，因此当前只能下载回答，无法在线提交。" : "The study database is not configured yet, so answers can be downloaded but not submitted online.");
  }
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function downloadAnswers() {
  const blob = new Blob([answerText()], { type: "text/plain;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${answers.response_code}.txt`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

async function submitResponse() {
  const config = window.__EVALUATION_CONFIG__ || {};
  if (!config.supabaseUrl || !config.supabaseAnonKey) {
    errorBox.textContent = copy[answers.language].submitError;
    return;
  }
  submitting = true;
  render();
  const payload = {
    response_code: answers.response_code,
    receipt_code: answers.receipt_code,
    questionnaire_version: VERSION,
    language: answers.language,
    device: answers.device || null,
    prior_use: answers.prior_use || null,
    quote_permission: answers.quote_permission || null,
    consent: Boolean(answers.consent),
    sus_responses: Object.fromEntries(Array.from({ length: 10 }, (_, index) => [`q${index + 1}`, answers[`sus_${index + 1}`] || null])),
    sus_score: calculateSus(),
    answers,
  };
  try {
    const endpoint = `${config.supabaseUrl.replace(/\/$/, "")}/rest/v1/evaluation_responses`;
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        apikey: config.supabaseAnonKey,
        Authorization: `Bearer ${config.supabaseAnonKey}`,
        "Content-Type": "application/json",
        Prefer: "return=minimal",
      },
      body: JSON.stringify(payload),
    });
    if (!response.ok) throw new Error(`Supabase returned ${response.status}`);
    submitted = true;
    localStorage.removeItem(STORAGE_KEY);
  } catch (error) {
    console.error("Evaluation submission failed", error);
    errorBox.textContent = copy[answers.language].submitError;
  } finally {
    submitting = false;
    render();
  }
}

backButton.addEventListener("click", () => {
  step = Math.max(0, step - 1);
  render();
});

nextButton.addEventListener("click", async () => {
  if (step === 0 && (!answers.eligible || !answers.consent)) {
    errorBox.textContent = copy[answers.language].requiredError;
    return;
  }
  if (step === 5) await submitResponse();
  else {
    step = Math.min(TOTAL_STEPS - 1, step + 1);
    render();
  }
});

languageSelect.addEventListener("change", () => {
  answers.language = languageSelect.value;
  saveDraft();
  render();
});

window.addEventListener("beforeunload", (event) => {
  if (!submitted && Object.keys(answers).length > 5) {
    event.preventDefault();
    event.returnValue = "";
  }
});

render();

