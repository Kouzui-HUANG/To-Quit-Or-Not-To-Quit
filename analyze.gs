// 問卷所有題目，依序對應到三大面向
var QUESTIONS = [
  '過去三個月內，我至少參加或自學一項新技能並已在工作中實際應用。',
  '每週至少有一項任務需要我嘗試之前未操作過的工具或方法。',
  '最近一次績效／任務回顧中，我的成果被具體點名為對團隊有明顯貢獻。',
  '在跨部門或跨專案合作中，我曾負責新領域的工作（過去三個月至少一次）。',
  '公司最近一次重大決策（如組織調整）前，員工能獲得充分資訊並理解自身影響。',
  '過去六個月內，公司曾因員工回饋而修訂流程或政策。',
  '在例行會議中，我提出的建議有超過一半獲得主管或團隊採納。',
  '公司對違反行為準則的人在過去三個月內採取過可見的處置。',
  '主管每月至少一次與我進行 1:1，並提供具體回饋與資源。',
  '過去三個月內，主管主動協調人力／預算來協助我完成目標。',
  '團隊成員曾於過去一個月內協作完成一項緊急任務而無衝突。',
  '過去三個月內，團隊中未出現公開羞辱或人身攻擊事件。',
  '最近四週，我平均每週睡眠不足六小時的夜晚不超過兩天。',
  '下班後的三個晚上，我能完全抽離工作並陪伴家人或休閒。',
  '過去一年內，醫生未因工作壓力對我提出健康警訊（如高血壓或焦慮）。',
  '最近一次健康檢查後，醫生認為我的壓力指標在可接受範圍。',
  '過去三個月，我成功使用至少一天特休／彈性假，而未被拒或要求改期。',
  '每週我能保留至少一天假期完全不處理工作訊息。',
  '過去三個月，我至少參加一次個人興趣或社交活動（如運動、課程、聚會）。',
  '過去兩年內，我的固定薪資或整體報酬至少調升一次且漲幅不低於產業平均。',
  '我擁有至少三個月的生活緊急預備金，可在無收入情況下維持基本開銷。',
  '公司提供的保險、分紅或獎金在過去一年內按時發放且無縮水。',
  '過去半年內，至少有一家獵頭或公司主動邀請我討論職缺。',
  '公司最新財報或公開說明中，所屬部門／產品線呈現成長或投入資源，而非縮編。',
  '我負責的專業技能在主要招聘網站的需求量於過去一年保持穩定或上升。'
];

// 各面向題目數量，用於計算平均
var CATEGORY_COUNTS = {
  work: 12,
  life: 7,
  reward: 6
};

function onFormSubmit(e) {
  var score = calculateScore(e.namedValues);
  var message = getResultMessage(score);
  var email = e.namedValues['若您希望在一週後收到您與其他填答者的比較分析，請留下 Email（選填）'];
  var stats = computeStatistics();
  updateSummarySheet(stats);

  if (email) {
    GmailApp.sendEmail(email, '離職自我檢測量表分析結果', '您的總分為 ' + score + ' 分。\n' + message);

    var trigger = ScriptApp.newTrigger('sendScheduledEmail')
      .timeBased()
      .after(7 * 24 * 60 * 60 * 1000)
      .create();
    PropertiesService.getScriptProperties().setProperty(
      trigger.getUniqueId(),
      JSON.stringify({ email: email, userAvg: score / QUESTIONS.length })
    );
  }
}

function calculateScore(namedValues) {
  var questions = QUESTIONS;
  var sum = 0;
  for (var i = 0; i < questions.length; i++) {
    var v = parseFloat(namedValues[questions[i]]);
    if (!isNaN(v)) sum += v;
  }
  return sum;
}

function getResultMessage(score) {
  if (score >= 80) {
    return '整體滿意度高，先思考如何加速成長或爭取資源。';
  } else if (score >= 60) {
    return '中等，專注改善低分面向，設定 3–6 個月觀察期。';
  } else if (score >= 40) {
    return '警戒，需要明確行動（內部提案或外部探索）。';
  } else {
    return '離職動機高，請同步準備履歷、財務緩衝與外部機會。';
  }
}

// === 統計計算與排程 ===

function computeStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses 1');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  var headers = data[0];
  var indices = QUESTIONS.map(function (q) {
    return headers.indexOf(q);
  });
  var workIdx = indices.slice(0, CATEGORY_COUNTS.work);
  var lifeIdx = indices.slice(CATEGORY_COUNTS.work, CATEGORY_COUNTS.work + CATEGORY_COUNTS.life);
  var rewardIdx = indices.slice(CATEGORY_COUNTS.work + CATEGORY_COUNTS.life);

  var totalArr = [];
  var workArr = [];
  var lifeArr = [];
  var rewardArr = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    var total = 0;
    var work = 0;
    var life = 0;
    var reward = 0;

    workIdx.forEach(function (i) {
      var v = parseFloat(row[i]);
      if (!isNaN(v)) work += v;
    });

    lifeIdx.forEach(function (i) {
      var v = parseFloat(row[i]);
      if (!isNaN(v)) life += v;
    });

    rewardIdx.forEach(function (i) {
      var v = parseFloat(row[i]);
      if (!isNaN(v)) reward += v;
    });

    total = work + life + reward;

    totalArr.push(total / QUESTIONS.length);
    workArr.push(work / CATEGORY_COUNTS.work);
    lifeArr.push(life / CATEGORY_COUNTS.life);
    rewardArr.push(reward / CATEGORY_COUNTS.reward);
  }

  return {
    total: { mean: mean(totalArr), median: median(totalArr) },
    work: { mean: mean(workArr), median: median(workArr) },
    life: { mean: mean(lifeArr), median: median(lifeArr) },
    reward: { mean: mean(rewardArr), median: median(rewardArr) }
  };
}

function mean(arr) {
  if (arr.length === 0) return 0;
  return arr.reduce(function (a, b) { return a + b; }, 0) / arr.length;
}

function median(arr) {
  if (arr.length === 0) return 0;
  var copy = arr.slice().sort(function (a, b) { return a - b; });
  var mid = Math.floor(copy.length / 2);
  if (copy.length % 2 === 0) {
    return (copy[mid - 1] + copy[mid]) / 2;
  } else {
    return copy[mid];
  }
}

function updateSummarySheet(stats) {
  if (!stats) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統計');
  if (!sheet) {
    sheet = ss.insertSheet('統計');
  } else {
    sheet.clear();
  }
  sheet.appendRow(['項目', '平均', '中位數']);
  sheet.appendRow(['總分', stats.total.mean, stats.total.median]);
  sheet.appendRow(['工作與環境', stats.work.mean, stats.work.median]);
  sheet.appendRow(['身心與生活', stats.life.mean, stats.life.median]);
  sheet.appendRow(['報酬與市場機會', stats.reward.mean, stats.reward.median]);
}

function sendScheduledEmail(e) {
  var props = PropertiesService.getScriptProperties();
  var id = e.triggerUid;
  var data = props.getProperty(id);
  if (!data) return;
  var info = JSON.parse(data);

  var stats = computeStatistics();
  var body = '您七天前填答的平均分數為 ' + info.userAvg.toFixed(2) + '。\n' +
             '目前所有填答者的平均分數為 ' + stats.total.mean.toFixed(2) + '。';
  GmailApp.sendEmail(info.email, '離職自我檢測量表比較結果', body);

  props.deleteProperty(id);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getUniqueId() === id) {
      ScriptApp.deleteTrigger(t);
    }
  });
}
