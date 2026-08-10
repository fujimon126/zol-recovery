/** 記録シートを読んで「グラフ」シートを作り直します。 */
function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = getSheet_();
  var addSheet = getAddSheet_();
  var dash = ss.getSheetByName(DASH_NAME);
  if (!dash) dash = ss.insertSheet(DASH_NAME);

  dash.getCharts().forEach(function (c) { dash.removeChart(c); });
  // clear() では結合が解けない。前回の結合が残ったまま表を書くとエラーになるので、
  // 先にシート全体の結合を解除する。
  dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).breakApart();
  dash.clear();

  var totalAdded = 0, totalRec = 0, totalCnt = 0, n = 0;
  var byMonth = {}, monthKeys = [];
  var byFactory = {}, factoryKeys = [];
  var byMachine = {}, machineKeys = [];

  function month_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
  function bucketMonth_(ym) {
    if (!byMonth[ym]) { byMonth[ym] = { added: 0, rec: 0, cnt: 0 }; monthKeys.push(ym); }
    return byMonth[ym];
  }
  function bucketFactory_(f) {
    if (!byFactory[f]) { byFactory[f] = { added: 0, rec: 0, cnt: 0 }; factoryKeys.push(f); }
    return byFactory[f];
  }

  // ドライ機への追加量
  var aLast = addSheet.getLastRow();
  if (aLast >= 2) {
    addSheet.getRange(2, 1, aLast - 1, ADD_HEADERS.length).getValues().forEach(function (v) {
      if (v[ACOL.ID] === '') return;
      var d = v[ACOL.TS] instanceof Date ? v[ACOL.TS] : new Date(v[ACOL.TS]);
      if (isNaN(d.getTime())) return;
      var amount = Number(v[ACOL.ADDED]) || 0;
      totalAdded += amount;
      bucketMonth_(month_(d)).added += amount;
      bucketFactory_(String(v[ACOL.FACTORY] || '')).added += amount;
    });
  }

  // 機械ごとの回収量と回転数
  var last = src.getLastRow();
  if (last >= 2) {
    src.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach(function (v) {
      if (v[COL.ID] === '') return;
      var d = v[COL.TS] instanceof Date ? v[COL.TS] : new Date(v[COL.TS]);
      if (isNaN(d.getTime())) return;
      var rec = Number(v[COL.REC]) || 0;
      var cnt = Number(v[COL.CNT]) || 0;
      var factory = String(v[COL.FACTORY] || '');
      totalRec += rec; totalCnt += cnt; n++;

      var m = bucketMonth_(month_(d));
      m.rec += rec; m.cnt += cnt;
      var f = bucketFactory_(factory);
      f.rec += rec; f.cnt += cnt;

      var mk = v[COL.MNAME] + ' No.' + v[COL.MNUM];
      if (!byMachine[mk]) { byMachine[mk] = { factory: factory, rec: 0, cnt: 0 }; machineKeys.push(mk); }
      byMachine[mk].rec += rec;
      byMachine[mk].cnt += cnt;
    });
  }

  if (!n && !totalAdded) {
    dash.getRange('A1').setValue('まだ記録がありません。アプリから記録するとここにグラフが出ます。')
      .setFontSize(12).setFontWeight('bold');
    return;
  }

  monthKeys.sort();
  factoryKeys.sort();
  var overallRate = rate_(totalRec, totalAdded);

  // ---------- 見出し ----------
  dash.getRange('A1').setValue('ゾール回収量 ダッシュボード')
    .setFontSize(18).setFontWeight('bold').setFontColor('#17201e');
  dash.getRange('A2').setValue(
    '最終更新 ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy年M月d日 HH:mm') +
    '  ／  記録が届くと自動で作り直されます'
  ).setFontColor('#77837f').setFontSize(10);

  dash.getRange('A3').setValue(
    '【回収率とは】 回収率 = 回収量 ÷ ドライ機への追加量。' +
    '追加したゾールのうち、どれだけ回収できたかを表します。数値が下がってきたら点検の目安です。'
  ).setFontSize(10).setFontColor('#4d5a56').setWrap(true);
  dash.setRowHeight(3, 32);

  // ---------- 全体 ----------
  var kpiHead = 5, kpiVal = 6;
  dash.getRange(kpiHead, 1, 1, 5)
    .setValues([['ドライ機への総追加量', '総回収量', '総回転数', '回収率', '記録件数']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45')
    .setHorizontalAlignment('center').setBorder(true, true, true, true, true, false, '#ccd4d0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(kpiVal, 1, 1, 5)
    .setValues([[totalAdded, totalRec, totalCnt, overallRate, n]])
    .setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, false, '#ccd4d0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(kpiVal, 1, 1, 2).setNumberFormat('#,##0.0" L"');
  dash.getRange(kpiVal, 3).setNumberFormat('#,##0" 回"');
  dash.getRange(kpiVal, 4).setNumberFormat('#,##0.0"%"').setFontColor(C_REC).setBackground('#fdf3e9');
  dash.getRange(kpiVal, 5).setNumberFormat('#,##0" 件"');
  dash.setRowHeight(kpiVal, 30);

  // ---------- 月別 ----------
  // グラフ用の列を隣どうしに並べておくと、範囲指定が単純で崩れにくい。
  var mTitle = 8, mHead = 9;
  sectionTitle_(dash, mTitle, '① 月別の推移');
  dash.getRange(mHead, 1, 1, 6)
    .setValues([['月', '追加量 (L)', '回収量 (L)', '回収率 (%)', '全期間の平均 (%)', '回転数 (回)']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  var mData = monthKeys.map(function (k) {
    var g = byMonth[k];
    return [k, g.added, g.rec, rate_(g.rec, g.added), overallRate, g.cnt];
  });
  if (mData.length) {
    dash.getRange(mHead + 1, 1, mData.length, 6).setValues(mData);
    dash.getRange(mHead + 1, 2, mData.length, 2).setNumberFormat('#,##0.0');
    dash.getRange(mHead + 1, 4, mData.length, 2).setNumberFormat('#,##0.0');
    dash.getRange(mHead + 1, 6, mData.length, 1).setNumberFormat('#,##0');
    banding_(dash, mHead + 1, mData.length, 6);
  }
  var mNote = mHead + Math.max(1, mData.length) + 1;
  readingNote_(dash, mNote,
    '【グラフの読み方】 灰色の棒がドライ機へ追加した量、オレンジの棒が回収できた量です。' +
    '緑の折れ線が回収率で、右側の目盛りで読みます。点線は全期間の平均なので、' +
    '折れ線が点線より下がった月は回収率が落ちています。');

  // ---------- 工場別 ----------
  var fTitle = mNote + 2, fHead = fTitle + 1;
  sectionTitle_(dash, fTitle, '② 工場別の集計');
  dash.getRange(fHead, 1, 1, 5)
    .setValues([['工場', '追加量 (L)', '回収量 (L)', '回収率 (%)', '回転数 (回)']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  var fData = factoryKeys.map(function (k) {
    var g = byFactory[k];
    return [k, g.added, g.rec, rate_(g.rec, g.added), g.cnt];
  });
  if (fData.length) {
    dash.getRange(fHead + 1, 1, fData.length, 5).setValues(fData);
    dash.getRange(fHead + 1, 2, fData.length, 3).setNumberFormat('#,##0.0');
    dash.getRange(fHead + 1, 5, fData.length, 1).setNumberFormat('#,##0');
    banding_(dash, fHead + 1, fData.length, 5);
  }
  var fNote = fHead + Math.max(1, fData.length) + 1;
  readingNote_(dash, fNote,
    '【グラフの読み方】 工場ごとに、追加した量と回収できた量を並べています。' +
    'オレンジ(回収量)が灰色(追加量)に近いほど回収できています。');

  // ---------- 機械別 ----------
  // 回収率の高い順に並べると、下位の機械がそのまま点検候補になる。
  machineKeys.sort(function (a, b) {
    var fa = byFactory[byMachine[a].factory] ? byFactory[byMachine[a].factory].added : 0;
    var fb = byFactory[byMachine[b].factory] ? byFactory[byMachine[b].factory].added : 0;
    return (rate_(byMachine[b].rec, fb) || 0) - (rate_(byMachine[a].rec, fa) || 0);
  });

  var xTitle = fNote + 2, xHead = xTitle + 1;
  sectionTitle_(dash, xTitle, '③ 機械別の集計');
  dash.getRange(xHead, 1, 1, 5)
    .setValues([['機械', '回収率 (%)', '回収量 (L)', '回転数 (回)', '工場']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  dash.getRange(xHead, 2).setNote(
    '回収率 = その機械の回収量 ÷ その工場の追加量。\n' +
    'ゾールはドライ機へまとめて追加するため、追加量は機械別に分けられません。\n' +
    '同じ工場の機械の回収率を足すと、その工場の回収率になります。');
  var xData = machineKeys.map(function (k) {
    var g = byMachine[k];
    var fa = byFactory[g.factory] ? byFactory[g.factory].added : 0;
    return [k, rate_(g.rec, fa), g.rec, g.cnt, g.factory];
  });
  if (xData.length) {
    dash.getRange(xHead + 1, 1, xData.length, 5).setValues(xData);
    dash.getRange(xHead + 1, 2, xData.length, 2).setNumberFormat('#,##0.0');
    dash.getRange(xHead + 1, 4, xData.length, 1).setNumberFormat('#,##0');
    banding_(dash, xHead + 1, xData.length, 5);
  }
  var xNote = xHead + Math.max(1, xData.length) + 1;
  readingNote_(dash, xNote,
    '【グラフの読み方】 回収率の高い順に並べています。棒が短い機械ほど回収できていません。' +
    '追加量は機械ごとに分けられないため、分母には「その機械が置かれている工場の追加量」を使っています。' +
    '同じ工場の機械の回収率を足すと、その工場の回収率になります。' +
    '円グラフは、回収量が機械ごとにどれくらいの割合かを示します。');

  // ---------- 体裁 ----------
  dash.setColumnWidth(1, 200);
  dash.setColumnWidth(2, 105);
  dash.setColumnWidth(3, 105);
  dash.setColumnWidth(4, 110);
  dash.setColumnWidth(5, 120);
  dash.setColumnWidth(6, 95);
  dash.setHiddenGridlines(true);

  // ---------- グラフ ----------
  var axisText = { color: '#4d5a56', fontSize: 11 };
  var titleStyle = { color: '#17201e', fontSize: 15, bold: true };
  var gridStyle = { color: '#e3e8e5' };

  // グラフは行に固定されるので、次のグラフの位置は高さから決めて重なりを防ぐ。
  var chartRow = 5;
  function place_(height) {
    var at = chartRow;
    chartRow += Math.ceil(height / 21) + 2;   // 既定の行の高さは約21px
    return at;
  }

  if (mData.length) {
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dash.getRange(mHead, 1, mData.length + 1, 5))
      .setPosition(place_(380), 8, 0, 0)
      .setOption('title', '① 月ごとに、追加した量・回収した量・回収率を見る')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 720).setOption('height', 380)
      .setOption('backgroundColor', { fill: '#ffffff', stroke: '#ccd4d0', strokeWidth: 1 })
      .setOption('chartArea', { left: 70, top: 70, width: '72%', height: '65%' })
      .setOption('legend', { position: 'top', alignment: 'center', textStyle: axisText })
      .setOption('bar', { groupWidth: '62%' })
      .setOption('series', {
        0: { type: 'bars', color: C_ADDED, targetAxisIndex: 0 },
        1: { type: 'bars', color: C_REC, targetAxisIndex: 0 },
        2: { type: 'line', color: C_RATE, targetAxisIndex: 1, lineWidth: 3, pointSize: 8, pointShape: 'circle' },
        3: { type: 'line', color: '#b0bdb8', targetAxisIndex: 1, lineWidth: 2, lineDashStyle: [6, 4], pointSize: 0 }
      })
      .setOption('vAxes', {
        0: { title: '追加量・回収量 (L)', titleTextStyle: axisText, textStyle: axisText, gridlines: gridStyle, viewWindow: { min: 0 } },
        1: { title: '回収率 (%)', titleTextStyle: axisText, textStyle: axisText, gridlines: { count: 0 }, viewWindow: { min: 0 } }
      })
      .setOption('hAxis', { title: '月', titleTextStyle: axisText, textStyle: axisText })
      .build());
  }

  if (fData.length) {
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(fHead, 1, fData.length + 1, 3))
      .setPosition(place_(330), 8, 0, 0)
      .setOption('title', '② 工場ごとの追加量と回収量をくらべる')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 720).setOption('height', 330)
      .setOption('backgroundColor', { fill: '#ffffff', stroke: '#ccd4d0', strokeWidth: 1 })
      .setOption('chartArea', { left: 70, top: 70, width: '80%', height: '62%' })
      .setOption('legend', { position: 'top', alignment: 'center', textStyle: axisText })
      .setOption('colors', [C_ADDED, C_REC])
      .setOption('bar', { groupWidth: '55%' })
      .setOption('vAxis', { title: 'リットル (L)', titleTextStyle: axisText, textStyle: axisText, gridlines: gridStyle, viewWindow: { min: 0 } })
      .setOption('hAxis', { textStyle: axisText })
      .build());
  }

  if (xData.length) {
    var barHeight = Math.max(280, 110 + xData.length * 44);
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dash.getRange(xHead, 1, xData.length + 1, 2))
      .setPosition(place_(barHeight), 8, 0, 0)
      .setOption('title', '③ 機械別の回収率(高い順)')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 720)
      .setOption('height', barHeight)
      .setOption('backgroundColor', { fill: '#ffffff', stroke: '#ccd4d0', strokeWidth: 1 })
      .setOption('chartArea', { left: 190, top: 60, width: '62%', height: '78%' })
      .setOption('legend', { position: 'none' })
      .setOption('colors', [C_RATE])
      .setOption('bar', { groupWidth: '62%' })
      .setOption('hAxis', { title: '回収率 (%)  ← 短いほど回収できていません', titleTextStyle: axisText, textStyle: axisText, gridlines: gridStyle, viewWindow: { min: 0 } })
      .setOption('vAxis', { textStyle: axisText })
      .build());

    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(xHead, 1, xData.length + 1, 1))
      .addRange(dash.getRange(xHead, 3, xData.length + 1, 1))
      .setPosition(place_(380), 8, 0, 0)
      .setOption('title', '④ 回収量の内訳(どの機械がどれだけ回収したか)')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 720).setOption('height', 380)
      .setOption('backgroundColor', { fill: '#ffffff', stroke: '#ccd4d0', strokeWidth: 1 })
      .setOption('chartArea', { left: 30, top: 70, width: '88%', height: '76%' })
      .setOption('pieHole', 0.45)
      .setOption('pieSliceText', 'percentage')
      .setOption('sliceVisibilityThreshold', 0)
      .setOption('legend', { position: 'right', textStyle: axisText })
      .build());
  }
}

/** 見出し行(左に色の帯を付けて区切りを分かりやすくします) */
function sectionTitle_(sheet, row, text) {
  sheet.getRange(row, 1, 1, 6).setBackground('#f6f8f7');
  sheet.getRange(row, 1).setValue(text)
    .setFontSize(13).setFontWeight('bold').setFontColor('#17201e');
  sheet.setRowHeight(row, 26);
}

/** 表の下に置く「読み方」の注釈 */
function readingNote_(sheet, row, text) {
  var range = sheet.getRange(row, 1, 1, 6);
  range.merge();
  range.setValue(text)
    .setFontSize(10).setFontColor('#4d5a56').setBackground('#fdf3e9')
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(row, 46);
}

/** 表を1行おきに薄く塗って読みやすくします */
function banding_(sheet, startRow, rows, cols) {
  for (var i = 0; i < rows; i++) {
    if (i % 2 === 1) sheet.getRange(startRow + i, 1, 1, cols).setBackground('#f6f8f7');
  }
  sheet.getRange(startRow, 1, rows, cols)
    .setBorder(null, null, true, null, null, true, '#e3e8e5', SpreadsheetApp.BorderStyle.SOLID);
}


