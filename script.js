const itineraryList = document.querySelector("#itinerary-list");
const flightList = document.querySelector("#flight-list");
const cityList = document.querySelector("#city-list");
const otherExpenseList = document.querySelector("#other-expense-list");
const addItineraryButton = document.querySelector("#add-itinerary");
const addFlightButton = document.querySelector("#add-flight");
const addCityButton = document.querySelector("#add-city");
const addOtherExpenseButton = document.querySelector("#add-other-expense");
const demoButton = document.querySelector("#load-demo");
const exportDraftButton = document.querySelector("#export-draft");
const importDraftButton = document.querySelector("#import-draft");
const draftFileInput = document.querySelector("#draft-file");
const restoreDraftWordButton = document.querySelector("#restore-draft-word");
const draftWordFileInput = document.querySelector("#draft-word-file");
const saveDataButton = document.querySelector("#save-data");
const loadDataButton = document.querySelector("#load-data");
const clearDataButton = document.querySelector("#clear-data");
const generateButton = document.querySelector("#generate-docx");
const statusText = document.querySelector("#status");
const topStatusText = document.querySelector("#top-status");
const rateDate = document.querySelector("#rate-date");
const documentDate = document.querySelector("#document-date");
const rateGrid = document.querySelector("#rate-grid");
const storageKey = "trip-docx-generator-data";
let autoSaveTimer = null;
let autoSaveReady = false;
let isRestoringForm = false;
const rateDefaults = {
  USD: "",
  EUR: "8",
  GBP: "9.13",
  JPY: "",
  HKD: "",
};
const rateValues = { ...rateDefaults };
const rateInputIds = {
  USD: "exchange-rate",
  EUR: "eur-rate",
  GBP: "gbp-rate",
  JPY: "jpy-rate",
  HKD: "hkd-rate",
};
const fallbackLocationStandards = [
  { type: "出访", country: "澳大利亚", city: "悉尼", currency: "USD", hotel: 200, meal: 60, misc: 50 },
  { type: "出访", country: "澳大利亚", city: "布里斯班", currency: "USD", hotel: 180, meal: 60, misc: 50 },
  { type: "出访", country: "美国", city: "华盛顿", currency: "USD", hotel: 240, meal: 55, misc: 45 },
  { type: "出访", country: "美国", city: "纽约", currency: "USD", hotel: 300, meal: 55, misc: 45 },
  { type: "出访", country: "英国", city: "伦敦", currency: "GBP", hotel: 180, meal: 45, misc: 35 },
  { type: "出访", country: "法国", city: "巴黎", currency: "EUR", hotel: 180, meal: 60, misc: 40 },
  { type: "出访", country: "德国", city: "柏林", currency: "EUR", hotel: 150, meal: 60, misc: 40 },
  { type: "出访", country: "日本", city: "东京", currency: "USD", hotel: 200, meal: 50, misc: 40 },
  { type: "出访", country: "韩国", city: "首尔", currency: "USD", hotel: 180, meal: 50, misc: 40 },
  { type: "出访", country: "新加坡", city: "新加坡", currency: "USD", hotel: 220, meal: 55, misc: 40 },
  { type: "培训", country: "中国", city: "北京", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "上海", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "广州", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "深圳", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "郑州", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "杭州", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "成都", currency: "CNY", hotel: "", meal: "", misc: "" },
  { type: "培训", country: "中国", city: "武汉", currency: "CNY", hotel: "", meal: "", misc: "" },
];
const locationStandards = globalThis.locationStandardsData || fallbackLocationStandards;
const defaultItinerary = [
  {
    date: "2026-07-20",
    content:
      "上午：上海-悉尼 宿悉尼\n12:35-15:00\n出关并取行李；之后从机场乘车前往悉尼市区（车程1小时）；\n午餐后，乘车前往新南威尔士州司法委员会做访问。",
  },
  {
    date: "2026-07-20",
    content:
      "下午：悉尼\n15:00-17:30\n访问新南威尔士州司法委员会\n与首席行政官、委员会成员等交流座谈。\n1.委员会的管理机制与组织结构。\n2.《1997年青少年罪犯法》所规定的警方警告、青少年司法会议及社区服务令的实际运行效果，以及法官在促进“交叉”（指不同司法体系或处理方式的交叉互动）方面的作用。\n3.委员会如何将关于儿童认知能力和心理创伤反应的专业课程纳入法官及地方法官的职前和在职培训。\n4.新南威尔士州司法系统在处理涉及土著青少年罪犯或受害者的案件时所适用的文化回应原则及具体司法指南。\n5.双方未来合作的潜力与可能性。\n17:30结束访问后，驱车前往酒店住宿。",
  },
  {
    date: "2026-07-21",
    content:
      "上午：悉尼\n9:00-10:30\n访问新南威尔士州最高法院\n与最高法院刑事庭法官、青少年司法官等交流座谈。\n交流内容：\n1.当前NSW在审理10至13岁未成年人案件时适用“Doli Incapax”（缺乏犯罪能力）原则的具体做法，以及该原则在近期立法改革中面临的挑战。\n2.在《保释法》修正案针对青少年设置“高门槛”保释条件的背景下，最高法院在审理青少年案件时平衡社区安全与“无罪推定”基本原则的司法立场与实践经验。\n3.在将刑事责任年龄从10岁提高至14岁的社会呼吁背景下，法院的现行立场，以及在法定年龄修改前通过证据规则（如Doli Incapax的严格适用）限制低龄儿童进入刑事司法系统的具体措施。\n4.针对土著青少年在司法系统中占比较高的问题，最高法院在审理相关案件时纳入考量的文化背景因素，以及“原住民专门法庭”模式在高院层面的价值与运作机制。\n11:00-12:30\n拜访新南威尔士州地方法院\n与青少年法官、法院治安官等交流座谈。\n交流内容：\n1.了解未成年人案件审理中的分案处理与法庭环境适幼化设计；\n2.了解涉罪未成年人的审前分流机制及其适用条件；\n3.探讨恢复性司法在青少年案件中的本土实践与成效评估；\n4.交流法院主导下的跨部门协作机制。",
  },
  {
    date: "2026-07-21",
    content:
      "下午：悉尼\n14:00-14:30\n乘车前往悉尼大学法学院（车程30分钟）\n14:30-17:00\n拜访悉尼大学法学院\n与院长威尔特、司法协作中心教授、青少年犯罪学教授等交流座谈。\n交流内容：\n1.了解澳大利亚在未成年人司法保护方面的法律体系、基本原则及主要制度异同。\n2.了解审前转处、社区矫正、家庭会议等非监禁措施的适用条件与效果评估。\n3.探讨被害人法律援助、心理干预、隐私保护及赔偿程序等实务机制。\n4.了解悉尼大学法学院在少年法庭设计、法官培训、团队协作等方面的经验。",
  },
  {
    date: "2026-07-21",
    content:
      "晚上：悉尼-布里斯班\n18:00-19:00 晚餐后，乘车前往机场（车程60分钟）",
  },
  {
    date: "2026-07-22",
    content:
      "上午：布里斯班\n9:00-10:30\n访问昆士兰州地方法院\n与地方法院治安法官、法院登记官、警方联络官等交流座谈。\n交流内容：\n1.青少年在首次出庭时须就指控作出认罪或不认罪的答辩，法院登记官和值班律师在该阶段向青少年及其家长说明程序权利和答辩后果的操作流程。\n2.对于谋杀、贩毒等严重犯罪，地方法院须先进行交付聆讯，审查控方证据是否充分足以将案件移交至更高级别法院（儿童法院或最高法院）的程序安排。\n3.治安法官在审理青少年保释申请时需纳入考量的特殊因素，包括青少年的年龄、心智成熟度、家庭监护能力以及社区支持系统的可用性。\n4.警方在启动正式诉讼程序前对青少年犯罪嫌疑人可采取的替代措施，包括不予处理、正式告诫及转介恢复性司法程序的层级化适用条件。",
  },
  {
    date: "2026-07-22",
    content:
      "10:40-12:00\n参加庭审\n1.观察法庭在审理未成年人案件时是否采用不同于成年人的程序安排。\n2.重点了解未成年人在庭审中是否独立获得法律援助、是否允许法定代理人或合适成年人全程陪同、是否享有专门的告知与解释程序。\n3.学习法庭在询问未成年人证人、被害人及被告人时是否采用特殊询问技巧或辅助工具（如视频作证、屏蔽设施等），以及如何确保证言真实性同时减少二次伤害。\n4.观察法官在宣判时是否对未成年人进行个别化教育与法律引导，是否明确解释判决理由、法律后果及未来的矫正或帮教措施。",
  },
  {
    date: "2026-07-22",
    content:
      "下午：布里斯班\n14:30-17:30\n访问昆士兰州儿童法院\n与儿童法院院长、儿童法院治安法官等交流座谈。\n1.了解贵法院的管理和运行机制。\n2.了解未成年人刑事案件的分流与转处机制。\n3.了解涉罪未成年人的社会调查与心理干预制度。\n4.学习澳大利亚未成年人被害人保护与救助体系。\n5.探讨双方未来可能的合作机制。\n17:30结束访问后，乘车前往住宿酒店。",
  },
  {
    date: "2026-07-23",
    content:
      "上午：布里斯班-上海\n抵达上海后，转机回郑州，圆满完成访问任务。",
  },
];

rateDate.value = toDateInput(new Date());
documentDate.value = toDateInput(new Date());
document.addEventListener("input", () => {
  updateComputedBudgetFields();
  scheduleAutoSave();
});
document.addEventListener("change", () => {
  renderRateInputs();
  updateComputedBudgetFields();
  scheduleAutoSave();
});
document.querySelector("#budget-type").addEventListener("change", () => {
  populateCountrySelect();
  syncCityRowsToBudgetType();
});
document.querySelector("#country").addEventListener("change", () => {
  syncCityRowsToCountry();
  ensureCityRowsForCountry();
  renderRateInputs();
});

addItineraryButton.addEventListener("click", () => {
  addItineraryRow("", "");
  scheduleAutoSave();
});
addFlightButton.addEventListener("click", () => {
  addFlightRow({});
  scheduleAutoSave();
});
addCityButton.addEventListener("click", () => {
  addCityRow({});
  scheduleAutoSave();
});
addOtherExpenseButton.addEventListener("click", () => {
  addOtherExpenseRow({});
  scheduleAutoSave();
});

demoButton.addEventListener("click", loadDemo);
exportDraftButton.addEventListener("click", exportDraft);
importDraftButton.addEventListener("click", () => draftFileInput.click());
draftFileInput.addEventListener("change", importDraft);
restoreDraftWordButton.addEventListener("click", () => draftWordFileInput.click());
draftWordFileInput.addEventListener("change", restoreDraftFromWord);
saveDataButton.addEventListener("click", saveCurrentData);
loadDataButton.addEventListener("click", loadSavedData);
clearDataButton.addEventListener("click", clearSavedData);
generateButton.addEventListener("click", generateDocx);

populateCountrySelect();
addItineraryRow("", "");
addFlightRow({});
addCityRow({});
addDefaultOtherExpenseRows();
renderRateInputs();
updateComputedBudgetFields();
autoSaveReady = true;
autoRestoreSavedData();

function loadDemo() {
  setStatus("正在填入示例...");
  setValue("#org-name", "某某单位");
  setValue("#leader-name", "张三");
  setValue("#budget-type", "出访");
  setValue("#approval-format", "henan");
  populateCountrySelect();
  setValue("#country", "澳大利亚");
  setValue("#people-count", "5");
  setValue("#exchange-rate", "6.82");
  setValue("#eur-rate", "8");
  setValue("#gbp-rate", "9.13");
  setValue("#jpy-rate", "0.047");
  setValue("#hkd-rate", "0.87");
  setValue("#document-date", "2026-05-04");
  Object.assign(rateValues, {
    USD: "6.82",
    EUR: "8",
    GBP: "9.13",
    JPY: "0.047",
    HKD: "0.87",
  });
  setValue("#business-flight-label", "国际大段公务舱+澳洲内陆段经济舱");
  setValue("#business-flight", "64000");
  setValue("#economy-flight", "26000");
  setValue("#business-count", "2");
  setValue("#economy-count", "3");

  itineraryList.innerHTML = "";
  defaultItinerary.forEach((item) => addItineraryRow(item.date, item.content));
  flightList.innerHTML = "";
  addFlightRow({
    date: "2026-07-20",
    departTime: "00:15",
    arriveTime: "12:35",
    code: "东方航空 MU735",
    from: "上海浦东国际机场 T1",
    to: "金斯福德史密斯机场 T1",
    note: "飞行10小时20分",
  });
  addFlightRow({
    date: "2026-07-21",
    departTime: "21:50",
    arriveTime: "23:25",
    code: "澳洲航空 QF558",
    from: "金斯福德史密斯机场 T3",
    to: "布里斯班机场 D",
    note: "飞行1小时35分",
  });
  addFlightRow({
    date: "2026-07-23",
    departTime: "11:15",
    arriveTime: "19:30",
    code: "东方航空 MU716",
    from: "布里斯班机场 国际航站楼",
    to: "上海浦东国际机场 T1",
    note: "飞行10小时15分",
  });
  cityList.innerHTML = "";
  addCityRow({ country: "澳大利亚", city: "悉尼", hotelNights: 1, mealDays: 2, miscDays: 2 });
  addCityRow({ country: "澳大利亚", city: "布里斯班", hotelNights: 2, mealDays: 2, miscDays: 2 });
  otherExpenseList.innerHTML = "";
  addOtherExpenseRow({ name: "城际交通", currency: "USD", amount: 1500, quantity: 4, quantityLabel: "天", splitByPeople: true });
  addOtherExpenseRow({ name: "购买境外保险", currency: "CNY", amount: 1200, quantity: 1, quantityLabel: "人", splitByPeople: false });
  addOtherExpenseRow({ name: "专业翻译费", currency: "USD", amount: 1200, quantity: 7, quantityLabel: "场", splitByPeople: true });
  addOtherExpenseRow({ name: "签证费", currency: "CNY", amount: "", quantity: 1, quantityLabel: "人", splitByPeople: false });
  renderRateInputs();
  updateComputedBudgetFields();
  scheduleAutoSave();
  setStatus("已填入示例。");
}

function addItineraryRow(date, content, insertAfter) {
  const row = document.createElement("div");
  row.className = "itinerary-row";
  row.innerHTML = `
    <label>
      日期
      <input class="itinerary-date" type="date" value="${escapeAttr(date)}">
    </label>
    <label>
      日程安排
      <textarea class="itinerary-content" placeholder="例如：上午访问某机构，下午交流座谈">${escapeHtml(content)}</textarea>
    </label>
    <div class="row-actions">
      <button class="ghost-button insert-itinerary-button" type="button">下方添加</button>
      <button class="delete-button" type="button">删除</button>
    </div>
  `;

  row.querySelector(".insert-itinerary-button").addEventListener("click", () => {
    const newRow = addItineraryRow("", "", row);
    newRow.querySelector(".itinerary-date").focus();
    scheduleAutoSave();
  });
  row.querySelector(".delete-button").addEventListener("click", () => {
    row.remove();
    scheduleAutoSave();
  });
  if (insertAfter instanceof HTMLElement) {
    insertAfter.after(row);
  } else {
    itineraryList.append(row);
  }
  return row;
}

function addFlightRow(data) {
  const row = document.createElement("div");
  row.className = "flight-row";
  row.innerHTML = `
    <label>
      日期
      <input class="flight-date" type="date">
    </label>
    <label>
      出发时间
      <input class="flight-depart-time" type="time">
    </label>
    <label>
      抵达时间
      <input class="flight-arrive-time" type="time">
    </label>
    <label>
      航班号/交通
      <input class="flight-code" type="text" placeholder="例如：MU735">
    </label>
    <label>
      出发地
      <input class="flight-from" type="text">
    </label>
    <label>
      抵达地
      <input class="flight-to" type="text">
    </label>
    <label>
      备注
      <input class="flight-note" type="text" placeholder="例如：飞行10小时20分">
    </label>
    <button class="delete-button" type="button">删除</button>
  `;

  row.querySelector(".flight-date").value = data.date || "";
  row.querySelector(".flight-depart-time").value = data.departTime || "";
  row.querySelector(".flight-arrive-time").value = data.arriveTime || "";
  row.querySelector(".flight-code").value = data.code || "";
  row.querySelector(".flight-from").value = data.from || "";
  row.querySelector(".flight-to").value = data.to || "";
  row.querySelector(".flight-note").value = data.note || "";
  row.querySelector(".delete-button").addEventListener("click", () => {
    row.remove();
    scheduleAutoSave();
  });
  flightList.append(row);
}

function addCityRow(data) {
  const country = data.country || value("#country") || firstCountryForCurrentType();
  const row = document.createElement("div");
  row.className = "city-row";
  row.innerHTML = `
    <label>
      国家/地区
      <select class="city-country"></select>
    </label>
    <label>
      省市/城市
      <select class="city-name"></select>
    </label>
    <label>
      可手动改城市
      <input class="city-custom" type="text" placeholder="自定义城市">
    </label>
    <label>
      币种
      <select class="city-currency">
        <option value="USD">美元</option>
        <option value="EUR">欧元</option>
        <option value="GBP">英镑</option>
        <option value="JPY">日元</option>
        <option value="HKD">港币</option>
        <option value="CNY">人民币</option>
      </select>
    </label>
    <label>
      住宿晚数
      <input class="hotel-nights" type="number" min="0" placeholder="0">
    </label>
    <label>
      住宿标准/晚
      <input class="hotel-rate" type="number" min="0">
    </label>
    <label>
      伙食天数
      <input class="meal-days" type="number" min="0" placeholder="0">
    </label>
    <label>
      伙食标准/天
      <input class="meal-rate" type="number" min="0">
    </label>
    <label>
      公杂天数
      <input class="misc-days" type="number" min="0" placeholder="0">
    </label>
    <label>
      公杂标准/天
      <input class="misc-rate" type="number" min="0">
    </label>
    <button class="delete-button" type="button">删除</button>
  `;

  cityList.append(row);
  populateRowCountrySelect(row, country);
  populateRowCitySelect(row, data.city);
  row.querySelector(".city-custom").value = data.customCity || "";
  row.querySelector(".hotel-nights").value = data.hotelNights ?? "";
  row.querySelector(".meal-days").value = data.mealDays ?? "";
  row.querySelector(".misc-days").value = data.miscDays ?? "";

  if (data.country || data.city) {
    applySelectedCityStandard(row, {
      keepManualRates: Boolean(data.hotelRate || data.mealRate || data.miscRate || data.currency),
    });
  }

  row.querySelector(".city-currency").value = data.currency || row.querySelector(".city-currency").value;
  row.querySelector(".hotel-rate").value = data.hotelRate ?? row.querySelector(".hotel-rate").value;
  row.querySelector(".meal-rate").value = data.mealRate ?? row.querySelector(".meal-rate").value;
  row.querySelector(".misc-rate").value = data.miscRate ?? row.querySelector(".misc-rate").value;

  row.querySelector(".city-country").addEventListener("change", () => {
    populateRowCitySelect(row);
    applySelectedCityStandard(row);
    renderRateInputs();
    updateComputedBudgetFields();
    scheduleAutoSave();
  });
  row.querySelector(".city-name").addEventListener("change", () => {
    applySelectedCityStandard(row);
    renderRateInputs();
    updateComputedBudgetFields();
    scheduleAutoSave();
  });
  row.querySelector(".delete-button").addEventListener("click", () => {
    row.remove();
    updateComputedBudgetFields();
    scheduleAutoSave();
  });
}

function addOtherExpenseRow(data) {
  const row = document.createElement("div");
  row.className = "other-expense-row";
  row.innerHTML = `
    <label>
      费用名称
      <input class="expense-name" type="text" placeholder="例如：签证费">
    </label>
    <label>
      币种
      <select class="expense-currency">
        <option value="USD">美元</option>
        <option value="EUR">欧元</option>
        <option value="GBP">英镑</option>
        <option value="JPY">日元</option>
        <option value="HKD">港币</option>
        <option value="CNY">人民币</option>
      </select>
    </label>
    <label>
      单价
      <input class="expense-amount" type="number" min="0">
    </label>
    <label>
      数量
      <input class="expense-quantity" type="number" min="0">
    </label>
    <label>
      单位
      <input class="expense-quantity-label" type="text" placeholder="天 / 场 / 人">
    </label>
    <label>
      分摊
      <select class="expense-split">
        <option value="yes">按出行人数分摊</option>
        <option value="no">不分摊，已是每人费用</option>
      </select>
    </label>
    <button class="delete-button" type="button">删除</button>
  `;

  row.querySelector(".expense-name").value = data.name || "";
  row.querySelector(".expense-currency").value = data.currency || "USD";
  row.querySelector(".expense-amount").value = data.amount ?? "";
  row.querySelector(".expense-quantity").value = data.quantity ?? "";
  row.querySelector(".expense-quantity-label").value = data.quantityLabel || "";
  row.querySelector(".expense-split").value = data.splitByPeople === false ? "no" : "yes";
  row.querySelector(".expense-currency").addEventListener("change", renderRateInputs);
  row.querySelector(".delete-button").addEventListener("click", () => {
    row.remove();
    renderRateInputs();
    updateComputedBudgetFields();
    scheduleAutoSave();
  });
  otherExpenseList.append(row);
}

function populateCountrySelect() {
  const countrySelect = document.querySelector("#country");
  const previous = countrySelect.value;
  const countries = countriesForCurrentType();

  countrySelect.innerHTML = countries.map((country) => `<option value="${escapeAttr(country)}">${escapeHtml(country)}</option>`).join("");
  countrySelect.value = countries.includes(previous) ? previous : countries[0] || "";
}

function populateRowCountrySelect(row, selectedCountry) {
  const select = row.querySelector(".city-country");
  const countries = countriesForCurrentType();

  select.innerHTML = countries.map((country) => `<option value="${escapeAttr(country)}">${escapeHtml(country)}</option>`).join("");
  select.value = countries.includes(selectedCountry) ? selectedCountry : countries[0] || "";
}

function populateRowCitySelect(row, selectedCity = "") {
  const country = row.querySelector(".city-country").value;
  const cities = locationStandards
    .filter((item) => item.type === value("#budget-type") && item.country === country)
    .sort((left, right) => sortChinese(left.city, right.city));
  const citySelect = row.querySelector(".city-name");
  citySelect.innerHTML = cities
    .map((item) => `<option value="${escapeAttr(item.city)}">${escapeHtml(item.city)}</option>`)
    .join("");
  citySelect.innerHTML = `<option value="">请选择</option>${citySelect.innerHTML}`;
  citySelect.value = cities.some((item) => item.city === selectedCity) ? selectedCity : cities[0]?.city || "";
  applySelectedCityStandard(row);
}

function applySelectedCityStandard(row, options = {}) {
  const standard = selectedCityStandard(row);
  if (!standard) return;

  row.querySelector(".city-currency").value = standard.currency;
  if (!options.keepManualRates) {
    row.querySelector(".hotel-rate").value = standard.hotel;
    row.querySelector(".meal-rate").value = standard.meal;
    row.querySelector(".misc-rate").value = standard.misc;
  }
}

function selectedCityStandard(row) {
  return locationStandards.find(
    (item) =>
      item.type === value("#budget-type") &&
      item.country === row.querySelector(".city-country").value &&
      item.city === row.querySelector(".city-name").value,
  );
}

function syncCityRowsToBudgetType() {
  [...document.querySelectorAll(".city-row")].forEach((row) => {
    populateRowCountrySelect(row, value("#country"));
    populateRowCitySelect(row);
    applySelectedCityStandard(row);
  });
  renderRateInputs();
  updateComputedBudgetFields();
}

function syncCityRowsToCountry() {
  [...document.querySelectorAll(".city-row")].forEach((row) => {
    row.querySelector(".city-country").value = value("#country");
    populateRowCitySelect(row);
    applySelectedCityStandard(row);
  });
  renderRateInputs();
  updateComputedBudgetFields();
}

function ensureCityRowsForCountry() {
  if (cityList.children.length === 0) {
    addCityRow({ country: value("#country") });
    return;
  }

  [...document.querySelectorAll(".city-row")].forEach((row) => {
    if (!row.querySelector(".city-name").value) {
      populateRowCitySelect(row);
      applySelectedCityStandard(row);
    }
  });
}

function saveCurrentData() {
  const payload = createFormPayload();
  localStorage.setItem(storageKey, JSON.stringify(payload));
  setStatus("已保存当前填写。下次打开这个网页，可以点击“加载上次填写”恢复。");
}

function scheduleAutoSave() {
  if (!autoSaveReady || isRestoringForm) return;
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(autoSaveCurrentData, 600);
}

function autoSaveCurrentData() {
  try {
    localStorage.setItem(storageKey, JSON.stringify(createFormPayload()));
    setStatus(`已自动保存 ${new Date().toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" })}`);
  } catch (error) {
    console.error(error);
    setStatus("自动保存失败，请点击“导出草稿”备份。");
  }
}

function createFormPayload() {
  updateComputedBudgetFields();
  return {
    app: "trip-docx-generator",
    version: 1,
    savedAt: new Date().toISOString(),
    fields: [...document.querySelectorAll("input, select")].reduce((fields, field) => {
      if (field.id && field.type !== "file") fields[field.id] = field.value;
      return fields;
    }, {}),
    itinerary: [...document.querySelectorAll(".itinerary-row")].map((row) => ({
      date: row.querySelector(".itinerary-date").value,
      content: row.querySelector(".itinerary-content").value,
    })),
    flights: [...document.querySelectorAll(".flight-row")].map((row) => ({
      date: row.querySelector(".flight-date").value,
      departTime: row.querySelector(".flight-depart-time").value,
      arriveTime: row.querySelector(".flight-arrive-time").value,
      code: row.querySelector(".flight-code").value,
      from: row.querySelector(".flight-from").value,
      to: row.querySelector(".flight-to").value,
      note: row.querySelector(".flight-note").value,
    })),
    cities: [...document.querySelectorAll(".city-row")].map((row) => ({
      country: row.querySelector(".city-country").value,
      city: row.querySelector(".city-name").value,
      customCity: row.querySelector(".city-custom").value,
      currency: row.querySelector(".city-currency").value,
      hotelNights: row.querySelector(".hotel-nights").value,
      hotelRate: row.querySelector(".hotel-rate").value,
      mealDays: row.querySelector(".meal-days").value,
      mealRate: row.querySelector(".meal-rate").value,
      miscDays: row.querySelector(".misc-days").value,
      miscRate: row.querySelector(".misc-rate").value,
    })),
    otherExpenses: [...document.querySelectorAll(".other-expense-row")].map((row) => ({
      name: row.querySelector(".expense-name").value,
      currency: row.querySelector(".expense-currency").value,
      amount: row.querySelector(".expense-amount").value,
      quantity: row.querySelector(".expense-quantity").value,
      quantityLabel: row.querySelector(".expense-quantity-label").value,
      splitByPeople: row.querySelector(".expense-split").value === "yes",
    })),
  };
}

function loadSavedData() {
  const saved = localStorage.getItem(storageKey);

  if (!saved) {
    setStatus("还没有保存过数据。");
    return;
  }

  try {
    const payload = JSON.parse(saved);
    restoreFormPayload(payload);
    setStatus("已加载上次保存的填写内容。");
  } catch {
    setStatus("保存的数据无法读取，可以清除后重新填写。");
  }
}

function autoRestoreSavedData() {
  const saved = localStorage.getItem(storageKey);
  if (!saved) return;

  try {
    restoreFormPayload(JSON.parse(saved));
    setStatus("已自动恢复上次填写内容。");
  } catch (error) {
    console.error(error);
    setStatus("上次自动保存的数据无法读取，可以清除后重新填写。");
  }
}

function restoreFormPayload(payload) {
  isRestoringForm = true;
  try {
    Object.entries(payload.fields || {}).forEach(([id, savedValue]) => {
      const field = document.getElementById(id);
      if (field) field.value = savedValue;
      const rateCurrency = Object.entries(rateInputIds).find(([, rateId]) => rateId === id)?.[0];
      if (rateCurrency) rateValues[rateCurrency] = savedValue;
    });
    populateCountrySelect();
    if (payload.fields?.country) setValue("#country", payload.fields.country);

    itineraryList.innerHTML = "";
    (payload.itinerary || []).forEach((item) => addItineraryRow(item.date, item.content));
    if (itineraryList.children.length === 0) addItineraryRow("", "");
    flightList.innerHTML = "";
    (payload.flights || []).forEach((item) => addFlightRow(item));
    if (flightList.children.length === 0) addFlightRow({});
    cityList.innerHTML = "";
    (payload.cities || []).forEach((item) => addCityRow(item));
    if (cityList.children.length === 0) addCityRow({});
    otherExpenseList.innerHTML = "";
    (payload.otherExpenses || []).forEach((item) => addOtherExpenseRow(item));
    if (otherExpenseList.children.length === 0) addOtherExpenseRow({});
    renderRateInputs();
    updateComputedBudgetFields();
  } finally {
    isRestoringForm = false;
  }
}

function clearSavedData() {
  localStorage.removeItem(storageKey);
  clearCurrentForm();
  setStatus("已清除保存数据，并清空当前页面。");
}

function exportDraft() {
  const payload = createFormPayload();
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  const country = value("#country") || "未命名";
  const date = toDateInput(new Date());
  link.href = URL.createObjectURL(blob);
  link.download = `出访日程及预算草稿-${country}-${date}.json`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
  setStatus("草稿已导出。下次可点击“导入草稿”完整恢复。");
}

async function importDraft(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;

  try {
    const payload = JSON.parse(await file.text());
    if (!payload || typeof payload !== "object" || !payload.fields) {
      throw new Error("Invalid draft file");
    }
    restoreFormPayload(payload);
    localStorage.setItem(storageKey, JSON.stringify(payload));
    setStatus("草稿已导入，页面内容已完整恢复。");
  } catch (error) {
    console.error(error);
    setStatus("草稿导入失败。请确认选择的是本工具导出的 .json 草稿文件。");
  }
}

async function restoreDraftFromWord(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;

  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const draftXml = await readZipText(bytes, "customXml/eaa-draft.xml");
    const xml = new DOMParser().parseFromString(draftXml, "application/xml");
    const jsonText = xml.querySelector("json")?.textContent || "";
    const payload = JSON.parse(jsonText);

    if (!payload || payload.app !== "trip-docx-generator" || !payload.fields) {
      throw new Error("Invalid embedded draft");
    }

    restoreFormPayload(payload);
    localStorage.setItem(storageKey, JSON.stringify(payload));
    setStatus("已从 Word 恢复草稿。这个功能只对新版工具生成的 Word 生效。");
  } catch (error) {
    console.error(error);
    setStatus("这个 Word 里没有可恢复的草稿。请使用新版工具生成的 Word，或导入 .json 草稿。");
  }
}

function generateDocx() {
  updateComputedBudgetFields();
  const draftPayload = createFormPayload();
  const data = collectData();
  const documentXml = buildDocumentXml(data);
  const files = buildDocxFiles(documentXml, draftPayload);
  const zipBytes = createZip(files);
  const blob = new Blob([zipBytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `出访总行程及预算-${data.country || "未命名"}.docx`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
  setStatus("Word 已生成。如果没有看到文件，请检查浏览器下载目录。");
}

function collectData() {
  const cityItems = collectCityItems();
  const otherExpenses = collectOtherExpenses();
  const hotelCny = round(cityItems.reduce((total, item) => total + item.hotelCny, 0));
  const mealCny = round(cityItems.reduce((total, item) => total + item.mealCny, 0));
  const miscCny = round(cityItems.reduce((total, item) => total + item.miscCny, 0));
  const stayMealMisc = hotelCny + mealCny + miscCny;
  const peopleCount = Math.max(numberValue("#people-count"), 1);
  const otherCny = round(otherExpenses.reduce((total, item) => total + item.cnyPerPerson, 0));
  const businessTotalPerPerson = numberValue("#business-flight") + stayMealMisc + otherCny;
  const economyTotalPerPerson = numberValue("#economy-flight") + stayMealMisc + otherCny;
  const totalBudget =
    businessTotalPerPerson * numberValue("#business-count") +
    economyTotalPerPerson * numberValue("#economy-count");

  return {
    orgName: value("#org-name"),
    leaderName: value("#leader-name"),
    country: value("#country"),
    budgetType: value("#budget-type"),
    approvalFormat: value("#approval-format"),
    peopleCount,
    rateDate: value("#rate-date"),
    documentDate: value("#document-date"),
    exchangeRate: numberValue("#exchange-rate"),
    eurRate: numberValue("#eur-rate"),
    gbpRate: numberValue("#gbp-rate"),
    jpyRate: numberValue("#jpy-rate"),
    hkdRate: numberValue("#hkd-rate"),
    businessFlight: numberValue("#business-flight"),
    economyFlight: numberValue("#economy-flight"),
    businessCount: numberValue("#business-count"),
    economyCount: numberValue("#economy-count"),
    cityItems,
    otherExpenses,
    hotelAmount: round(cityItems.reduce((total, item) => total + item.hotelOriginal, 0)),
    mealAmount: round(cityItems.reduce((total, item) => total + item.mealOriginal, 0)),
    miscAmount: round(cityItems.reduce((total, item) => total + item.miscOriginal, 0)),
    otherCny,
    usedCurrencies: usedCurrencies(cityItems, otherExpenses),
    hotelCny,
    mealCny,
    miscCny,
    stayMealMisc,
    businessFlightLabel: value("#business-flight-label") || "公务舱机票",
    businessTotalPerPerson,
    economyTotalPerPerson,
    totalBudget,
    itinerary: [...document.querySelectorAll(".itinerary-row")].map((row) => ({
      date: row.querySelector(".itinerary-date").value,
      content: row.querySelector(".itinerary-content").value.trim(),
    })),
    flights: collectFlights(),
  };
}

function collectFlights() {
  return [...document.querySelectorAll(".flight-row")]
    .map((row) => ({
      date: row.querySelector(".flight-date").value,
      departTime: row.querySelector(".flight-depart-time").value,
      arriveTime: row.querySelector(".flight-arrive-time").value,
      code: row.querySelector(".flight-code").value.trim(),
      from: row.querySelector(".flight-from").value.trim(),
      to: row.querySelector(".flight-to").value.trim(),
      note: row.querySelector(".flight-note").value.trim(),
    }))
    .filter((item) => item.date || item.departTime || item.arriveTime || item.code || item.from || item.to || item.note);
}

function buildDocumentXml(data) {
  if (data.approvalFormat === "hainan") {
    return buildHainanDocumentXml(data);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraph("出访总行程", "Title")}
    ${paragraph(`组团单位盖章：${data.orgName}                                    团长审核签名：${data.leaderName}`, "Normal")}
    ${itineraryTable(data.itinerary)}
    ${paragraph(`${data.budgetType === "培训" ? "赴" : "访问"}${data.country}经费预算`, "BudgetTitle")}
    ${paragraph(`依据《财行[2013]516号》、《财行[2017]434号》文件，贵院${data.peopleCount}人${data.budgetType === "培训" ? "赴" : "出访"}${data.country}境外费用明细如下，访问结束后根据实际结算多退少补。`)}
    ${paragraph("一、国际旅费", "Normal")}
    ${paragraph(`${data.businessFlightLabel}：${money(data.businessFlight)}元/人； `)}
    ${paragraph(`全程经济舱：${money(data.economyFlight)} 元/人；`)}
    ${paragraph("机票价格实时变动，以出票价格为准。")}
    ${paragraph(`${formatShortDate(data.rateDate)}汇率：`)}
    ${currencyRateParagraphs(data.usedCurrencies)}
    ${paragraph("二、住宿费：")}
    ${cityBudgetParagraphs(data.cityItems, "hotel")}
    ${paragraph(`住宿费小计：${money(data.hotelCny)}元人民币/人`)}
    ${paragraph("三、伙食费：")}
    ${cityBudgetParagraphs(data.cityItems, "meal")}
    ${paragraph(`伙食费小计：${money(data.mealCny)}元人民币/人`)}
    ${paragraph("四、公杂费：")}
    ${cityBudgetParagraphs(data.cityItems, "misc")}
    ${paragraph(`公杂费小计：${money(data.miscCny)}元人民币/人`)}
    ${paragraph("住宿+伙食+公杂小计：")}
    ${paragraph(`${money(data.hotelCny)} + ${money(data.mealCny)} + ${money(data.miscCny)} = ${money(data.stayMealMisc)}元人民币/人`)}
    ${paragraph("五、其他费用：")}
    ${otherExpenseParagraphs(data.otherExpenses, data.peopleCount)}
    ${paragraph(`其他费用小计：${otherExpenseSubtotalText(data.otherExpenses)}= ${money(data.otherCny)}元人民币/人`)}
    ${paragraph("六、预算总计：")}
    ${paragraph(`公务舱总计：${money(data.businessFlight)} + ${money(data.stayMealMisc)} + ${money(data.otherCny)} = ${money(data.businessTotalPerPerson)}元人民币/人`)}
    ${paragraph(`经济舱总计：${money(data.economyFlight)} + ${money(data.stayMealMisc)} + ${money(data.otherCny)} = ${money(data.economyTotalPerPerson)}元人民币/人`)}
    ${paragraph("总预算：")}
    ${paragraph(`公务舱${data.businessCount}人+经济舱${data.economyCount}人=${money(data.businessTotalPerPerson * data.businessCount)} + ${money(data.economyTotalPerPerson * data.economyCount)}= ${money(data.totalBudget)}元人民币`)}
    ${paragraph("", "BudgetBlank")}
    ${paragraph("", "BudgetBlank")}
    ${paragraph(formatShortDate(data.documentDate), "BudgetDate")}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="850" w:right="1474" w:bottom="850" w:left="1474" w:header="851" w:footer="1559" w:gutter="0"/>
      <w:pgNumType w:fmt="numberInDash"/>
      <w:cols w:space="0" w:num="1"/>
      <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function buildHainanDocumentXml(data) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraph("出访计划日程", "HainanTitle")}
    ${hainanItineraryTable(data.itinerary, data.flights)}
    ${paragraph(`${data.budgetType === "培训" ? "赴" : "访问"}${data.country}经费预算`, "HainanBudgetTitle")}
    ${budgetParagraphs(data)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
      <w:cols w:space="425" w:num="1"/>
      <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function budgetParagraphs(data) {
  return `
    ${paragraph(`依据《财行[2013]516号》、《财行[2017]434号》文件，贵院${data.peopleCount}人${data.budgetType === "培训" ? "赴" : "出访"}${data.country}境外费用明细如下，访问结束后根据实际结算多退少补。`)}
    ${paragraph("一、国际旅费", "Normal")}
    ${paragraph(`${data.businessFlightLabel}：${money(data.businessFlight)}元/人； `)}
    ${paragraph(`全程经济舱：${money(data.economyFlight)} 元/人；`)}
    ${paragraph("机票价格实时变动，以出票价格为准。")}
    ${paragraph(`${formatShortDate(data.rateDate)}汇率：`)}
    ${currencyRateParagraphs(data.usedCurrencies)}
    ${paragraph("二、住宿费：")}
    ${cityBudgetParagraphs(data.cityItems, "hotel")}
    ${paragraph(`住宿费小计：${money(data.hotelCny)}元人民币/人`)}
    ${paragraph("三、伙食费：")}
    ${cityBudgetParagraphs(data.cityItems, "meal")}
    ${paragraph(`伙食费小计：${money(data.mealCny)}元人民币/人`)}
    ${paragraph("四、公杂费：")}
    ${cityBudgetParagraphs(data.cityItems, "misc")}
    ${paragraph(`公杂费小计：${money(data.miscCny)}元人民币/人`)}
    ${paragraph("住宿+伙食+公杂小计：")}
    ${paragraph(`${money(data.hotelCny)} + ${money(data.mealCny)} + ${money(data.miscCny)} = ${money(data.stayMealMisc)}元人民币/人`)}
    ${paragraph("五、其他费用：")}
    ${otherExpenseParagraphs(data.otherExpenses, data.peopleCount)}
    ${paragraph(`其他费用小计：${otherExpenseSubtotalText(data.otherExpenses)}= ${money(data.otherCny)}元人民币/人`)}
    ${paragraph("六、预算总计：")}
    ${paragraph(`公务舱总计：${money(data.businessFlight)} + ${money(data.stayMealMisc)} + ${money(data.otherCny)} = ${money(data.businessTotalPerPerson)}元人民币/人`)}
    ${paragraph(`经济舱总计：${money(data.economyFlight)} + ${money(data.stayMealMisc)} + ${money(data.otherCny)} = ${money(data.economyTotalPerPerson)}元人民币/人`)}
    ${paragraph("总预算：")}
    ${paragraph(`公务舱${data.businessCount}人+经济舱${data.economyCount}人=${money(data.businessTotalPerPerson * data.businessCount)} + ${money(data.economyTotalPerPerson * data.economyCount)}= ${money(data.totalBudget)}元人民币`)}
    ${paragraph("", "BudgetBlank")}
    ${paragraph("", "BudgetBlank")}
    ${paragraph(formatShortDate(data.documentDate), "BudgetDate")}
  `;
}

function itineraryTable(items) {
  const rows = items
    .filter((item) => item.date || item.content)
    .map((item) => tableRow([formatDate(item.date), item.content || ""]))
    .join("");

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="9"/>
        <w:tblW w:w="10306" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblBorders>
          <w:top w:val="single" w:color="000000" w:sz="4" w:space="0"/>
          <w:left w:val="single" w:color="000000" w:sz="4" w:space="0"/>
          <w:bottom w:val="single" w:color="000000" w:sz="4" w:space="0"/>
          <w:right w:val="single" w:color="000000" w:sz="4" w:space="0"/>
          <w:insideH w:val="single" w:color="000000" w:sz="4" w:space="0"/>
          <w:insideV w:val="single" w:color="000000" w:sz="4" w:space="0"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
        <w:tblCellMar>
          <w:top w:w="0" w:type="dxa"/>
          <w:left w:w="108" w:type="dxa"/>
          <w:bottom w:w="0" w:type="dxa"/>
          <w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2060"/><w:gridCol w:w="8246"/></w:tblGrid>
      ${tableRow(["日期", "日程安排"], true)}
      ${rows}
      ${noteRow()}
    </w:tbl>`;
}

function hainanItineraryTable(items, flights = []) {
  const grouped = [];
  items
    .filter((item) => item.date || item.content)
    .forEach((item) => {
      const last = grouped[grouped.length - 1];
      if (last && last.date === item.date) {
        last.content = `${last.content}\n${item.content || ""}`.trim();
      } else {
        grouped.push({ ...item });
      }
    });
  flights
    .filter((flight) => flight.date && !grouped.some((item) => item.date === flight.date))
    .forEach((flight) => grouped.push({ date: flight.date, content: "" }));
  const rows = grouped
    .sort((left, right) => String(left.date || "").localeCompare(String(right.date || "")))
    .map((item, index) => hainanDayRows(item, index + 1, flights.filter((flight) => flight.date === item.date)))
    .join("");

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="11"/>
        <w:tblW w:w="9704" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblBorders>
          <w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>
          <w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/>
          <w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>
          <w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/>
          <w:insideH w:val="single" w:color="auto" w:sz="4" w:space="0"/>
          <w:insideV w:val="single" w:color="auto" w:sz="4" w:space="0"/>
        </w:tblBorders>
        <w:tblLayout w:type="autofit"/>
        <w:tblCellMar>
          <w:top w:w="0" w:type="dxa"/>
          <w:left w:w="108" w:type="dxa"/>
          <w:bottom w:w="0" w:type="dxa"/>
          <w:right w:w="108" w:type="dxa"/>
        </w:tblCellMar>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="1756"/><w:gridCol w:w="7948"/></w:tblGrid>
      ${rows}
    </w:tbl>`;
}

function hainanDayRows(item, dayNumber, flightsForDay = []) {
  const lines = String(item.content || "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const heading = lines[0] && !isTimeLine(lines[0]) ? lines.shift() : "";
  const dayTitle = `第${chineseDayNumber(dayNumber)}天${formatDate(item.date)}${heading ? heading.replace(/^上午：|^下午：|^晚上：/, "") : ""}`;
  const rows = [`<w:tr><w:trPr><w:trHeight w:val="90" w:hRule="atLeast"/><w:jc w:val="center"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9704" w:type="dxa"/><w:gridSpan w:val="2"/><w:vAlign w:val="center"/></w:tcPr>${paragraph(dayTitle, "HainanDayTitle")}</w:tc></w:tr>`];
  const detailRows = flightsForDay.map(flightToHainanRow);

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const timeMatch = splitTimeLine(line);
    if (timeMatch) {
      const details = timeMatch.rest ? [timeMatch.rest] : [];
      while (lines[index + 1] && !isTimeLine(lines[index + 1])) {
        details.push(lines[index + 1]);
        index += 1;
      }
      detailRows.push({ time: timeMatch.time, content: details.join("\n") });
    } else {
      detailRows.push({ time: "", content: line });
    }
  }

  detailRows
    .sort((left, right) => timeSortValue(left.time) - timeSortValue(right.time))
    .forEach((row) => rows.push(hainanTableRow(row.time, row.content)));

  if (rows.length === 1) rows.push(hainanTableRow("", item.content || ""));
  return rows.join("");
}

function flightToHainanRow(flight) {
  const time = [flight.departTime, flight.arriveTime].filter(Boolean).join("-");
  const code = flight.code || "";
  const vehicle = code && /^乘坐/.test(code) ? code : code ? `乘坐${code}` : "";
  const route = [flight.from, flight.to].filter(Boolean).join(" - ");
  const note = flight.note ? `（${flight.note}）` : "";
  return {
    time,
    content: [vehicle, route, note].filter(Boolean).join("\n"),
  };
}

function timeSortValue(value) {
  const match = String(value || "").match(/^(\d{1,2})[:：](\d{2})/);
  if (!match) return 9999;
  return Number(match[1]) * 60 + Number(match[2]);
}

function hainanTableRow(time, content) {
  return `<w:tr><w:trPr><w:trHeight w:val="503" w:hRule="atLeast"/><w:jc w:val="center"/></w:trPr>${hainanCell(time, 1756, "HainanTimeText")}${hainanCell(content, 7948, "HainanTableText")}</w:tr>`;
}

function hainanCell(text, width, style) {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${String(text || "")
    .split("\n")
    .map((line) => paragraph(line, style))
    .join("")}</w:tc>`;
}

function isTimeLine(value) {
  return /^\d{1,2}[:：]\d{2}\s*[-–—]\s*\d{1,2}[:：]\d{2}/.test(value);
}

function splitTimeLine(value) {
  const match = String(value || "").match(/^(\d{1,2}[:：]\d{2}\s*[-–—]\s*\d{1,2}[:：]\d{2})(.*)$/);
  if (!match) return null;
  return {
    time: match[1].trim(),
    rest: match[2].trim(),
  };
}

function chineseDayNumber(number) {
  return ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"][number] || String(number);
}

function cityBudgetParagraphs(items, kind) {
  const config = {
    hotel: { days: "hotelNights", rate: "hotelRate", total: "hotelOriginal", cny: "hotelCny", unit: "晚" },
    meal: { days: "mealDays", rate: "mealRate", total: "mealOriginal", cny: "mealCny", unit: "天" },
    misc: { days: "miscDays", rate: "miscRate", total: "miscOriginal", cny: "miscCny", unit: "天" },
  }[kind];

  return items
    .filter((item) => item[config.days] || item[config.rate] || item[config.total])
    .map((item) => {
      const originalTotal = `${money(item[config.total])}${currencyName(item.currency)}/人`;
      const cnyTotal = `${money(item[config.cny])}元人民币/人`;
      const conversion = item.currency === "CNY" ? cnyTotal : `${originalTotal}*${rateText(item.currency)}=${cnyTotal}`;
      return paragraph(
        `${item.city}：${item[config.days]}${config.unit}*${money(item[config.rate])}${currencyName(
          item.currency,
        )}/人=${conversion}`,
      );
    })
    .join("");
}

function otherExpenseParagraphs(items, peopleCount) {
  return activeOtherExpenses(items)
    .map((item) => {
      const unit = item.quantityLabel ? item.quantityLabel : "";
      const rate = item.currency === "CNY" ? "" : `*${rateText(item.currency)}`;
      if (!item.splitByPeople) {
        return paragraph(
          `${item.name || "其他费用"}：${money(item.amount)}${currencyName(item.currency)}/人${rate}=${money(
            item.cnyPerPerson,
          )}元人民币/人`,
        );
      }
      return paragraph(
        `${item.name || "其他费用"}：${money(item.amount)}${currencyName(item.currency)}/${unit || "项"}*${item.quantity || 0}${unit}${rate}／${peopleCount}人=${money(item.cnyPerPerson)}元人民币/人`,
      );
    })
    .join("");
}

function activeOtherExpenses(items) {
  return items.filter((item) => Number(item.amount) > 0 && Number(item.cnyPerPerson) > 0);
}

function otherExpenseSubtotalText(items) {
  const parts = activeOtherExpenses(items).map((item) => money(item.cnyPerPerson));
  return parts.length ? parts.join("+") : "0";
}

function currencyRateParagraphs(currencies) {
  return currencies
    .filter((currency) => currency !== "CNY" && Number(rateText(currency)) > 0)
    .map((currency) => paragraph(`1${currencyName(currency)}=${rateText(currency)}人民币（元）`))
    .join("");
}

function currencyName(currency) {
  return {
    USD: "美元",
    EUR: "欧元",
    GBP: "英镑",
    JPY: "日元",
    HKD: "港币",
    CNY: "元人民币",
  }[currency];
}

function rateText(currency) {
  return {
    USD: numberValue("#exchange-rate"),
    EUR: numberValue("#eur-rate"),
    GBP: numberValue("#gbp-rate"),
    JPY: numberValue("#jpy-rate"),
    HKD: numberValue("#hkd-rate"),
    CNY: "1",
  }[currency];
}

function tableRow(cells, isHeader = false) {
  return `<w:tr>${cells.map((cell, index) => tableCell(cell, index === 0 ? 2060 : 8246, isHeader)).join("")}</w:tr>`;
}

function tableCell(text, width, isHeader) {
  return `<w:tc>
    <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>
    ${String(text || "")
      .split("\n")
      .map((line) => paragraph(line, isHeader ? "TableHeader" : "TableText"))
      .join("")}
  </w:tc>`;
}

function noteRow() {
  const note =
    "附注：1.须注明国际和国外城市间交通工具及离、抵时间。\n2.未经外事审批部门批准，严禁变更出访路线、延长在国（境）外停留时间。\n3.公务活动要具体详实，若签署协议，须在日程中体现。";
  return `<w:tr><w:tc><w:tcPr><w:tcW w:w="10306" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>${note
    .split("\n")
    .map((line) => paragraph(line, "TableText"))
    .join("")}</w:tc></w:tr>`;
}

function paragraph(text, style = "Normal") {
  const styleMap = {
    Title: {
      pPr: '<w:jc w:val="center"/>',
      rPr: runPr("华文中宋", 44, true),
    },
    BudgetTitle: {
      pPr: '<w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/>',
      rPr: runPr("宋体", 30, true),
    },
    HainanTitle: {
      pPr: '<w:jc w:val="center"/>',
      rPr: runPr("方正小标宋简体", 48, false),
    },
    HainanBudgetTitle: {
      pPr: '<w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/>',
      rPr: runPr("宋体", 28, true),
    },
    HainanTableBold: {
      pPr: '<w:jc w:val="center"/><w:spacing w:line="440" w:lineRule="exact"/>',
      rPr: runPr("宋体", 28, true),
    },
    HainanDayTitle: {
      pPr: '<w:jc w:val="left"/><w:spacing w:line="440" w:lineRule="exact"/>',
      rPr: runPr("仿宋", 28, true),
    },
    HainanTimeText: {
      pPr: '<w:jc w:val="center"/><w:spacing w:line="440" w:lineRule="exact"/>',
      rPr: runPr("仿宋", 28),
    },
    HainanTableText: {
      pPr: '<w:jc w:val="left"/><w:spacing w:line="440" w:lineRule="exact"/>',
      rPr: runPr("仿宋", 28),
    },
    TableHeader: {
      pPr: '<w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/>',
      rPr: runPr("仿宋_GB2312", 24),
    },
    TableText: {
      pPr: '<w:jc w:val="both"/><w:spacing w:line="240" w:lineRule="auto"/>',
      rPr: runPr("仿宋_GB2312", 24),
    },
    BudgetDate: {
      pPr: '<w:jc w:val="right"/><w:spacing w:line="280" w:lineRule="auto"/>',
      rPr: runPr("仿宋_GB2312", 28),
    },
    BudgetBlank: {
      pPr: '<w:jc w:val="right"/><w:spacing w:line="280" w:lineRule="auto"/>',
      rPr: runPr("仿宋_GB2312", 28),
    },
    Normal: {
      pPr: '<w:jc w:val="both"/><w:spacing w:line="280" w:lineRule="auto"/>',
      rPr: runPr("仿宋_GB2312", 28),
    },
  };
  const config = styleMap[style] || styleMap.Normal;

  return `<w:p><w:pPr>${config.pPr}<w:rPr>${config.rPr}</w:rPr></w:pPr><w:r><w:rPr>${config.rPr}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function buildDocxFiles(documentXml, draftPayload) {
  return [
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/customXml/eaa-draft.xml" ContentType="application/xml"/>
</Types>`,
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    },
    {
      name: "word/document.xml",
      content: documentXml,
    },
    {
      name: "word/_rels/document.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    },
    {
      name: "word/styles.xml",
      content: stylesXml(),
    },
    {
      name: "customXml/eaa-draft.xml",
      content: draftPayloadXml(draftPayload),
    },
  ];
}

function draftPayloadXml(payload) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<eaaDraft xmlns="urn:eaa-trip-draft">
  <json>${escapeXml(JSON.stringify(payload))}</json>
</eaaDraft>`;
}

async function readZipText(bytes, targetName) {
  const entry = zipEntries(bytes).find((item) => item.name === targetName);
  if (!entry) throw new Error(`${targetName} not found`);
  const content = bytes.slice(entry.offset, entry.offset + entry.compressedSize);
  const uncompressed =
    entry.method === 0 ? content : entry.method === 8 ? await inflateZipContent(content) : (() => {
      throw new Error(`Unsupported compression method: ${entry.method}`);
    })();
  return new TextDecoder("utf-8").decode(uncompressed);
}

function zipEntries(bytes) {
  const end = findEndOfCentralDirectory(bytes);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const entryCount = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true);
  const decoder = new TextDecoder("utf-8");
  const entries = [];

  for (let index = 0; index < entryCount; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("Invalid central directory");
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.slice(offset + 46, offset + 46 + nameLength));
    const localNameLength = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
    entries.push({
      name,
      method,
      compressedSize,
      offset: localHeaderOffset + 30 + localNameLength + localExtraLength,
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(bytes) {
  for (let index = bytes.length - 22; index >= 0; index -= 1) {
    if (bytes[index] === 0x50 && bytes[index + 1] === 0x4b && bytes[index + 2] === 0x05 && bytes[index + 3] === 0x06) {
      return index;
    }
  }
  throw new Error("End of central directory not found");
}

async function inflateZipContent(bytes) {
  if (!("DecompressionStream" in window)) throw new Error("DecompressionStream is not supported");
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${style("Normal", "paragraph", "仿宋_GB2312", 28)}
  ${style("Title", "paragraph", "华文中宋", 44, true)}
  ${style("BudgetTitle", "paragraph", "宋体", 30, true)}
  ${style("HainanTitle", "paragraph", "方正小标宋简体", 48)}
  ${style("HainanBudgetTitle", "paragraph", "宋体", 28, true)}
  ${style("HainanTableBold", "paragraph", "宋体", 28, true)}
  ${style("HainanDayTitle", "paragraph", "仿宋", 28, true)}
  ${style("HainanTimeText", "paragraph", "仿宋", 28)}
  ${style("HainanTableText", "paragraph", "仿宋", 28)}
  ${style("TableHeader", "paragraph", "仿宋_GB2312", 24)}
  ${style("TableText", "paragraph", "仿宋_GB2312", 24)}
  ${style("BudgetDate", "paragraph", "仿宋_GB2312", 28)}
  ${style("BudgetBlank", "paragraph", "仿宋_GB2312", 28)}
</w:styles>`;
}

function style(id, type, font, size, bold = false) {
  return `<w:style w:type="${type}" w:styleId="${id}">
    <w:name w:val="${id}"/>
    <w:rPr>
      ${bold ? "<w:b/>" : ""}
      <w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}"/>
      <w:sz w:val="${size}"/>
      <w:szCs w:val="${size}"/>
    </w:rPr>
  </w:style>`;
}

function runPr(font, size, bold = false) {
  return `<w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="${font}" w:cs="Times New Roman"/>${
    bold ? "<w:b/><w:bCs/>" : ""
  }<w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>`;
}

function createZip(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const contentBytes = encoder.encode(file.content);
    const crc = crc32(contentBytes);
    const localHeader = concatBytes(
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(contentBytes.length),
      u32(contentBytes.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes,
    );

    localParts.push(localHeader, contentBytes);

    const centralHeader = concatBytes(
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(contentBytes.length),
      u32(contentBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      nameBytes,
    );

    centralParts.push(centralHeader);
    offset += localHeader.length + contentBytes.length;
  });

  const centralDirectory = concatBytes(...centralParts);
  const localDirectory = concatBytes(...localParts);
  const end = concatBytes(
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(centralDirectory.length),
    u32(localDirectory.length),
    u16(0),
  );

  return concatBytes(localDirectory, centralDirectory, end);
}

function crc32(bytes) {
  let crc = -1;
  for (const byte of bytes) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

const crcTable = Array.from({ length: 256 }, (_, index) => {
  let c = index;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return c >>> 0;
});

function u16(value) {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff]);
}

function u32(value) {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]);
}

function concatBytes(...arrays) {
  const length = arrays.reduce((sum, array) => sum + array.length, 0);
  const output = new Uint8Array(length);
  let offset = 0;
  arrays.forEach((array) => {
    output.set(array, offset);
    offset += array.length;
  });
  return output;
}

function value(selector) {
  return document.querySelector(selector).value.trim();
}

function numberValue(selector) {
  const field = document.querySelector(selector);
  return Number(field?.value || 0) || 0;
}

function setValue(selector, newValue) {
  const field = document.querySelector(selector);
  if (field) field.value = newValue;
}

function setStatus(message) {
  statusText.textContent = message;
  topStatusText.textContent = message;
}

function clearCurrentForm() {
  [...document.querySelectorAll("input")].forEach((input) => {
    if (input.type === "date") {
      input.value = input.id === "rate-date" || input.id === "document-date" ? toDateInput(new Date()) : "";
    } else if (!input.readOnly) {
      input.value = "";
    }
  });
  itineraryList.innerHTML = "";
  flightList.innerHTML = "";
  cityList.innerHTML = "";
  otherExpenseList.innerHTML = "";
  populateCountrySelect();
  addItineraryRow("", "");
  addFlightRow({});
  addCityRow({});
  addDefaultOtherExpenseRows();
  Object.assign(rateValues, rateDefaults);
  renderRateInputs();
  updateComputedBudgetFields();
}

function addDefaultOtherExpenseRows() {
  addOtherExpenseRow({ name: "城际交通", currency: "USD", quantityLabel: "天", splitByPeople: true });
  addOtherExpenseRow({ name: "购买境外保险", currency: "CNY", quantityLabel: "人", splitByPeople: false });
  addOtherExpenseRow({ name: "专业翻译费", currency: "USD", quantityLabel: "场", splitByPeople: true });
  addOtherExpenseRow({ name: "签证费", currency: "CNY", quantityLabel: "人", splitByPeople: false });
}

function unique(values) {
  return [...new Set(values)];
}

function sortChinese(left, right) {
  return String(left || "").localeCompare(String(right || ""), "zh-Hans-CN", { sensitivity: "base" });
}

function countriesForCurrentType() {
  return unique(
    locationStandards
      .filter((item) => item.type === value("#budget-type"))
      .map((item) => item.country),
  ).sort(sortChinese);
}

function firstCountryForCurrentType() {
  return countriesForCurrentType()[0] || "";
}

function renderRateInputs() {
  Object.entries(rateInputIds).forEach(([currency, id]) => {
    const field = document.getElementById(id);
    if (field) rateValues[currency] = field.value;
  });
  const existingValues = Object.fromEntries(
    Object.entries(rateInputIds).map(([currency]) => [currency, rateValues[currency] ?? rateDefaults[currency] ?? ""]),
  );
  const currencies = activeCurrencies();

  rateGrid.innerHTML = currencies
    .filter((currency) => currency !== "CNY")
    .map((currency) => {
      const step = currency === "JPY" || currency === "HKD" ? "0.0001" : "0.01";
      const value = existingValues[currency] || rateDefaults[currency] || "";
      return `
        <label>
          1 ${currencyName(currency)} = 人民币
          <input id="${rateInputIds[currency]}" type="number" min="0" step="${step}" value="${escapeAttr(value)}" placeholder="填写今日汇率">
        </label>
      `;
    })
    .join("");

  Object.entries(rateInputIds).forEach(([currency, id]) => {
    const field = document.getElementById(id);
    if (field) {
      field.addEventListener("input", () => {
        rateValues[currency] = field.value;
      });
    }
  });
}

function activeCurrencies() {
  const countryCurrency = locationStandards.find(
    (item) => item.type === value("#budget-type") && item.country === value("#country") && item.currency !== "CNY",
  )?.currency;
  return unique([countryCurrency, ...usedCurrencies(collectCityItems(), collectOtherExpenses())].filter(Boolean));
}

function updateComputedBudgetFields() {
  const cityItems = collectCityItems();
  const otherExpenses = collectOtherExpenses();
  const peopleCount = Math.max(numberValue("#people-count"), 1);
  const hotelAmount = round(cityItems.reduce((total, item) => total + item.hotelOriginal, 0));
  const mealAmount = round(cityItems.reduce((total, item) => total + item.mealOriginal, 0));
  const miscAmount = round(cityItems.reduce((total, item) => total + item.miscOriginal, 0));
  const otherCny = round(otherExpenses.reduce((total, item) => total + item.cnyPerPerson, 0));

  setValue("#hotel-amount", hotelAmount || "");
  setValue("#meal-amount", mealAmount || "");
  setValue("#misc-amount", miscAmount || "");
  setValue("#other-cny", otherCny);
}

function collectOtherExpenses() {
  const peopleCount = Math.max(numberValue("#people-count"), 1);
  return [...document.querySelectorAll(".other-expense-row")].map((row) => {
    const amount = numberFromRow(row, ".expense-amount");
    const quantity = numberFromRow(row, ".expense-quantity");
    const currency = row.querySelector(".expense-currency").value;
    const splitByPeople = row.querySelector(".expense-split").value === "yes";
    const total = splitByPeople ? amount * quantity : amount;
    const cnyPerPerson = round(convertToCny(total, currency) / (splitByPeople ? peopleCount : 1));

    return {
      name: row.querySelector(".expense-name").value.trim(),
      currency,
      amount,
      quantity,
      quantityLabel: row.querySelector(".expense-quantity-label").value.trim(),
      splitByPeople,
      total,
      cnyPerPerson,
    };
  });
}

function usedCurrencies(cityItems, otherExpenses) {
  return unique([
    ...cityItems
      .filter((item) => item.hotelOriginal || item.mealOriginal || item.miscOriginal)
      .map((item) => item.currency),
    ...otherExpenses.filter((item) => item.total).map((item) => item.currency),
  ]);
}

function collectCityItems() {
  return [...document.querySelectorAll(".city-row")].map((row) => {
    const currency = row.querySelector(".city-currency").value;
    const cityName = row.querySelector(".city-custom").value.trim() || row.querySelector(".city-name").value;
    const hotelOriginal = numberFromRow(row, ".hotel-nights") * numberFromRow(row, ".hotel-rate");
    const mealOriginal = numberFromRow(row, ".meal-days") * numberFromRow(row, ".meal-rate");
    const miscOriginal = numberFromRow(row, ".misc-days") * numberFromRow(row, ".misc-rate");

    return {
      country: row.querySelector(".city-country").value,
      city: cityName,
      currency,
      hotelNights: numberFromRow(row, ".hotel-nights"),
      hotelRate: numberFromRow(row, ".hotel-rate"),
      hotelOriginal,
      hotelCny: convertToCny(hotelOriginal, currency),
      mealDays: numberFromRow(row, ".meal-days"),
      mealRate: numberFromRow(row, ".meal-rate"),
      mealOriginal,
      mealCny: convertToCny(mealOriginal, currency),
      miscDays: numberFromRow(row, ".misc-days"),
      miscRate: numberFromRow(row, ".misc-rate"),
      miscOriginal,
      miscCny: convertToCny(miscOriginal, currency),
    };
  });
}

function convertToCny(amount, currency) {
  const rates = {
    USD: numberValue("#exchange-rate"),
    EUR: numberValue("#eur-rate"),
    GBP: numberValue("#gbp-rate"),
    JPY: numberValue("#jpy-rate"),
    HKD: numberValue("#hkd-rate"),
    CNY: 1,
  };
  return amount * (rates[currency] || 0);
}

function numberFromRow(row, selector) {
  return Number(row.querySelector(selector).value) || 0;
}

function toDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(`${value}T00:00:00`);
  const week = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"][date.getDay()];
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日（${week}）`;
}

function formatShortDate(value) {
  if (!value) return "";
  const date = new Date(`${value}T00:00:00`);
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`;
}

function money(value) {
  return String(Math.round(value));
}

function round(value) {
  return Math.round(value);
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}
