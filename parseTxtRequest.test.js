const assert = require('assert');

global.PropertiesService = {
  getScriptProperties() {
    return {
      getProperty: () => '',
      setProperty: () => {}
    };
  }
};

function offsetMinutes(date, timeZone) {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  const parts = dtf.formatToParts(date).reduce((acc, cur) => {
    acc[cur.type] = cur.value;
    return acc;
  }, {});
  const asUtc = Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day),
    Number(parts.hour),
    Number(parts.minute),
    Number(parts.second)
  );
  return (asUtc - date.getTime()) / 60000;
}

global.Utilities = {
  Charset: { SHIFT_JIS: 'Shift_JIS' },
  formatDate(date, timeZone, format) {
    if (format === 'Z') {
      const mins = offsetMinutes(date, timeZone);
      const sign = mins >= 0 ? '+' : '-';
      const abs = Math.abs(mins);
      const hh = String(Math.floor(abs / 60)).padStart(2, '0');
      const mm = String(abs % 60).padStart(2, '0');
      return `${sign}${hh}${mm}`;
    }
    if (format === "yyyy-MM-dd'T'HH:mm:ss'Z'") {
      return new Date(date.getTime()).toISOString().replace('.000Z', 'Z');
    }
    return new Date(date.getTime()).toISOString();
  },
  newBlob: (bytes, mimeType, name) => ({
    getDataAsString: () => '',
    setContentType: () => {},
    copyBlob: () => ({
      setContentType: () => {},
      getDataAsString: () => ''
    })
  })
};

global.LockService = {
  getScriptLock: () => ({
    tryLock: () => true,
    releaseLock: () => {}
  })
};

global.Gmail = {
  Users: {
    Labels: {
      list: () => ({ labels: [] }),
      get: () => null,
      create: () => ({ id: 'dummy' })
    },
    Threads: { modify: () => {} }
  }
};

global.MailApp = { sendEmail: () => {} };
global.GASunzip = { unzip: () => [] };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '' }) };

const fs = require('fs');
const vm = require('vm');
vm.runInThisContext(fs.readFileSync(require.resolve('./g2s.js'), 'utf8'));

const sampleTxt = `
[査定依頼日時・査定依頼番号] 2024年5月1日10時30分(123456)
商品：査定
ブランド名：トヨタ
車種名：プリウス
年式：2020
グレード：S
ボディタイプ・カテゴリ：ハッチバック
車体色・ドア数：ホワイトパールクリスタルシャイン／5ドア
ハンドル：右
燃料：ハイブリッド
ミッション：AT
駆動方式：FF
排気量：1800cc
走行距離：20000km
車検時期：2025年3月
事故歴：なし
クルマの状態・ラベル：良好
売却希望時期・ラベル：3か月以内
型式・装備：DAA-ZVW50, ナビ/ETC
その他オプション：サンルーフ
ご依頼者名：山田太郎様
ご依頼者カナ名：ヤマダタロウ様
郵便番号：123-4567
ご住所：東京都渋谷区1-2-3
メールアドレス：test@example.com
電話番号：090-1234-5678
その他の連絡先：03-1111-2222
連絡可能時間帯：午前中
`;

const parsed = parseTxtRequest_(sampleTxt, 'Asia/Tokyo');

assert.strictEqual(parsed.bodyColor, 'ホワイトパールクリスタルシャイン');
assert.strictEqual(parsed.doorCount, '5');
assert.strictEqual(parsed.modelCode, 'DAA-ZVW50');
assert.strictEqual(parsed.equipmentInfo, 'ナビ/ETC');
assert.strictEqual(parsed.desiredSellTiming, '3か月以内');
assert.strictEqual(parsed.carCondition, '良好');
assert.strictEqual(parsed.bodyType, 'ハッチバック');
assert.strictEqual(parsed.otherOptions, 'サンルーフ');
assert.strictEqual(parsed.customerName, '山田太郎');
assert.strictEqual(parsed.customerKana, 'ヤマダタロウ');
assert.strictEqual(parsed.requestDateIso, "2024-05-01T01:30:00Z");

console.log('parseTxtRequest_ sample parsing ✅');
