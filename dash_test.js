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
    return { records: 0, adds: 0, charts: 0 };
  }

  monthKeys.sort();
  factoryKeys.sort();
  var overallRate = rate_(totalRec, totalAdded);

  function machineRate_(mk) {
    var g = byMachine[mk];
    var fa = byFactory[g.factory] ? byFactory[g.factory].added : 0;
    return rate_(g.rec, fa);
  }
  // 工場ごとにまとめ、そのなかで回収率の高い順。点検すべき機械が下に集まる。
  machineKeys.sort(function (a, b) {
    if (byMachine[a].factory !== byMachine[b].factory) {
      return byMachine[a].factory < byMachine[b].factory ? -1 : 1;
    }
    return (machineRate_(b) || 0) - (machineRate_(a) || 0);
  });

  var axisText = { color: '#4d5a56', fontSize: 11 };
  var titleStyle = { color: '#17201e', fontSize: 15, bold: true };
  var gridStyle = { color: '#e3e8e5' };
  var panel = { fill: '#ffffff', stroke: '#ccd4d0', strokeWidth: 1 };

  // 表は左、グラフは右。機械が多いと表が広がるのでグラフ位置をずらす。
  var chartCol = Math.max(8, machineKeys.length + 3);
  var chartRow = 5;
  function place_(height) {
    var at = chartRow;
    chartRow += Math.ceil(height / 21) + 2;
    return at;
  }

  // ---------- 見出し ----------
  dash.getRange('A1').setValue('ゾール回収量 ダッシュボード')
    .setFontSize(18).setFontWeight('bold').setFontColor('#17201e');
  dash.getRange('A2').setValue(
    '最終更新 ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy年M月d日 HH:mm') +
    '  ／  記録が届くと自動で作り直されます'
  ).setFontColor('#77837f').setFontSize(10);
  dash.getRange('A3').setValue(
    '【回収率とは】 回収率 = 回収量 ÷ ドライ機への追加量。追加したゾールのうち、どれだけ回収できたかを表します。'
  ).setFontSize(10).setFontColor('#4d5a56');

  // ---------- 全体 ----------
  dash.getRange(5, 1, 1, 5)
    .setValues([['ドライ機への総追加量', '総回収量', '総回転数', '全体の回収率', '記録件数']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45')
    .setHorizontalAlignment('center');
  dash.getRange(6, 1, 1, 5)
    .setValues([[totalAdded, totalRec, totalCnt, overallRate, n]])
    .setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  dash.getRange(6, 1, 1, 3).setNumberFormat('#,##0.0" L"');
  dash.getRange(6, 3).setNumberFormat('#,##0" 回"');
  dash.getRange(6, 4).setNumberFormat('#,##0.0"%"').setFontColor(C_REC).setBackground('#fdf3e9');
  dash.getRange(6, 5).setNumberFormat('#,##0" 件"');
  dash.setRowHeight(6, 32);

  // ---------- ① 工場別の回収率 ----------
  var fTitle = 8, fHead = 9;
  sectionTitle_(dash, fTitle, '① 工場ごとの回収率');
  dash.getRange(fHead, 1, 1, 5)
    .setValues([['工場', '回収率 (%)', '追加量 (L)', '回収量 (L)', '回転数 (回)']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  var fData = factoryKeys.map(function (k) {
    var g = byFactory[k];
    return [k, rate_(g.rec, g.added), g.added, g.rec, g.cnt];
  });
  if (fData.length) {
    dash.getRange(fHead + 1, 1, fData.length, 5).setValues(fData);
    dash.getRange(fHead + 1, 2, fData.length, 3).setNumberFormat('#,##0.0');
    dash.getRange(fHead + 1, 5, fData.length, 1).setNumberFormat('#,##0');
    // 回収率の列を大きく見せる
    dash.getRange(fHead + 1, 2, fData.length, 1)
      .setFontSize(13).setFontWeight('bold').setFontColor(C_REC).setBackground('#fdf3e9');
    banding_(dash, fHead + 1, fData.length, 5);
  }
  var fNote = fHead + Math.max(1, fData.length) + 1;
  readingNote_(dash, fNote,
    '【見方】 棒が長い工場ほど、追加したゾールを多く回収できています。' +
    '棒の色の内訳は、その工場のどの機械が回収した分かを示します。' +
    '内訳を足すと、その工場の回収率になります。');

  // ---------- ② 機械別の回収率 ----------
  var xTitle = fNote + 2, xHead = xTitle + 1;
  sectionTitle_(dash, xTitle, '② 機械ごとの回収率');
  dash.getRange(xHead, 1, 1, 5)
    .setValues([['機械', '回収率 (%)', '工場', '回収量 (L)', '回転数 (回)']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  dash.getRange(xHead, 2).setNote(
    '回収率 = その機械の回収量 ÷ その工場の追加量。\n' +
    'ゾールはドライ機へまとめて追加するため、追加量は機械別に分けられません。\n' +
    '同じ工場の機械の回収率を足すと、その工場の回収率になります。');
  var xData = machineKeys.map(function (k) {
    var g = byMachine[k];
    return [k, machineRate_(k), g.factory, g.rec, g.cnt];
  });
  if (xData.length) {
    dash.getRange(xHead + 1, 1, xData.length, 5).setValues(xData);
    dash.getRange(xHead + 1, 2, xData.length, 1)
      .setNumberFormat('#,##0.0').setFontSize(12).setFontWeight('bold').setFontColor(C_RATE).setBackground('#e9f3ec');
    dash.getRange(xHead + 1, 4, xData.length, 1).setNumberFormat('#,##0.0');
    dash.getRange(xHead + 1, 5, xData.length, 1).setNumberFormat('#,##0');
    banding_(dash, xHead + 1, xData.length, 5);
  }
  var xNote = xHead + Math.max(1, xData.length) + 1;
  readingNote_(dash, xNote,
    '【見方】 工場ごとにまとめ、そのなかで回収率の高い順に並べています。' +
    '下にある機械ほど回収できていないので、点検の候補になります。');

  // ---------- ③ 工場×機械の内訳(積み上げ用の表) ----------
  // 1本の棒＝1工場。棒の長さがその工場の回収率、色の内訳が機械ごとの回収率。
  var sTitle = xNote + 2, sHead = sTitle + 1;
  sectionTitle_(dash, sTitle, '③ 内訳の計算表(グラフ用)');
  var stackHeader = ['工場'].concat(machineKeys);
  dash.getRange(sHead, 1, 1, stackHeader.length).setValues([stackHeader])
    .setFontWeight('bold').setFontSize(9).setBackground('#e3e8e5').setFontColor('#3c4a45');
  var sData = factoryKeys.map(function (f) {
    var row = [f];
    machineKeys.forEach(function (mk) {
      row.push(byMachine[mk].factory === f ? (machineRate_(mk) || 0) : 0);
    });
    return row;
  });
  if (sData.length) {
    dash.getRange(sHead + 1, 1, sData.length, stackHeader.length).setValues(sData);
    dash.getRange(sHead + 1, 2, sData.length, machineKeys.length).setNumberFormat('#,##0.0');
  }
  var sNote = sHead + Math.max(1, sData.length) + 1;
  readingNote_(dash, sNote, '※ ①のグラフを作るための表です。数字は「その機械が工場の回収率に占める分」です。');

  // ---------- ④ 月別 ----------
  var mTitle = sNote + 2, mHead = mTitle + 1;
  sectionTitle_(dash, mTitle, '④ 月別の推移');
  dash.getRange(mHead, 1, 1, 6)
    .setValues([['月', '追加量 (L)', '回収量 (L)', '回収率 (%)', '全期間の平均 (%)', '回転数 (回)']])
    .setFontWeight('bold').setFontSize(10).setBackground('#e3e8e5').setFontColor('#3c4a45');
  var mData = monthKeys.map(function (k) {
    var g = byMonth[k];
    return [k, g.added, g.rec, rate_(g.rec, g.added), overallRate, g.cnt];
  });
  if (mData.length) {
    dash.getRange(mHead + 1, 1, mData.length, 6).setValues(mData);
    dash.getRange(mHead + 1, 2, mData.length, 4).setNumberFormat('#,##0.0');
    dash.getRange(mHead + 1, 6, mData.length, 1).setNumberFormat('#,##0');
    banding_(dash, mHead + 1, mData.length, 6);
  }
  var mNote = mHead + Math.max(1, mData.length) + 1;
  readingNote_(dash, mNote,
    '【見方】 灰色の棒が追加した量、オレンジの棒が回収できた量です。緑の折れ線が回収率(右の目盛り)。' +
    '点線は全期間の平均なので、折れ線が点線より下がった月は回収率が落ちています。');

  // ---------- 体裁 ----------
  dash.setColumnWidth(1, 210);
  dash.setColumnWidth(2, 110);
  dash.setColumnWidth(3, 120);
  dash.setColumnWidth(4, 105);
  dash.setColumnWidth(5, 100);
  dash.setColumnWidth(6, 95);
  dash.setHiddenGridlines(true);

  // ---------- グラフ ----------
  var made = 0;

  // ① 工場ごとの回収率(機械の内訳つき積み上げ)
  if (sData.length) {
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dash.getRange(sHead, 1, sData.length + 1, stackHeader.length))
      .setPosition(place_(Math.max(300, 130 + sData.length * 70)), chartCol, 0, 0)
      .setOption('title', '① 工場ごとの回収率(色は機械ごとの内訳)')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 760)
      .setOption('height', Math.max(300, 130 + sData.length * 70))
      .setOption('backgroundColor', panel)
      .setOption('chartArea', { left: 130, top: 60, width: '58%', height: '70%' })
      .setOption('isStacked', true)
      .setOption('legend', { position: 'right', textStyle: { color: '#4d5a56', fontSize: 10 } })
      .setOption('bar', { groupWidth: '55%' })
      .setOption('hAxis', { title: '回収率 (%)  ← 棒が長いほど回収できています', titleTextStyle: axisText, textStyle: axisText, gridlines: gridStyle, viewWindow: { min: 0 } })
      .setOption('vAxis', { textStyle: { color: '#17201e', fontSize: 12, bold: true } })
      .build());
    made++;
  }

  // ② 機械ごとの回収率
  if (xData.length) {
    var barHeight = Math.max(300, 120 + xData.length * 46);
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dash.getRange(xHead, 1, xData.length + 1, 2))
      .setPosition(place_(barHeight), chartCol, 0, 0)
      .setOption('title', '② 機械ごとの回収率(下ほど回収できていません)')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 760)
      .setOption('height', barHeight)
      .setOption('backgroundColor', panel)
      .setOption('chartArea', { left: 210, top: 60, width: '60%', height: '80%' })
      .setOption('legend', { position: 'none' })
      .setOption('colors', [C_RATE])
      .setOption('bar', { groupWidth: '62%' })
      .setOption('hAxis', { title: '回収率 (%)', titleTextStyle: axisText, textStyle: axisText, gridlines: gridStyle, viewWindow: { min: 0 } })
      .setOption('vAxis', { textStyle: { color: '#17201e', fontSize: 11 } })
      .build());
    made++;
  }

  // ③ 月別の推移
  if (mData.length) {
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dash.getRange(mHead, 1, mData.length + 1, 5))
      .setPosition(place_(380), chartCol, 0, 0)
      .setOption('title', '③ 月ごとの追加量・回収量と回収率')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 760).setOption('height', 380)
      .setOption('backgroundColor', panel)
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
    made++;
  }

  // ④ 回収量の内訳
  if (xData.length) {
    dash.insertChart(dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(xHead, 1, xData.length + 1, 1))
      .addRange(dash.getRange(xHead, 4, xData.length + 1, 1))
      .setPosition(place_(380), chartCol, 0, 0)
      .setOption('title', '④ 回収量の内訳(どの機械がどれだけ回収したか)')
      .setOption('titleTextStyle', titleStyle)
      .setOption('width', 760).setOption('height', 380)
      .setOption('backgroundColor', panel)
      .setOption('chartArea', { left: 30, top: 70, width: '88%', height: '76%' })
      .setOption('pieHole', 0.45)
      .setOption('pieSliceText', 'percentage')
      .setOption('sliceVisibilityThreshold', 0)
      .setOption('legend', { position: 'right', textStyle: axisText })
      .build());
    made++;
  }

  return { records: n, adds: aLast >= 2 ? aLast - 1 : 0, charts: made };
}

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


/** 毎日午前6時に自動でグラフを更新します。 */
