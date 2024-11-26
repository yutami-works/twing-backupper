// モジュール
const fs = require('node:fs');
const fsp = fs.promises;
const path = require('node:path');
const https = require('node:https');
const xlsx = require('xlsx');

// 設定
const backupFolder = './backup'; // バックアップフォルダのパス
const excelFile = 'db.xlsx'; // 読み込むExcelファイル

// 画像をダウンロードして保存する関数
const downloadImage = async (url, savePath) => {
  return new Promise((resolve, reject) => {
    https.get(url, (response) => {
      if (response.statusCode !== 200) {
        reject(new Error(`HTTP Status Code: ${response.statusCode}`));
        return;
      }

      const fileStream = fs.createWriteStream(savePath);
      response.pipe(fileStream);
      fileStream.on('finish', () => fileStream.close(resolve));
    }).on('error', (err) => {
      reject(err);
    });
  });
}

// ファイルが存在するか確認する関数
const fileExists = async (filePath) => {
  try {
    await fsp.access(filePath);
    return true;
  } catch {
    return false;
  }
}

// メイン処理
const processExcel = async (filePath) => {
  const workbook = xlsx.readFile(filePath);
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const records = xlsx.utils.sheet_to_json(sheet);

    for (const record of records) {
      const { date, hash } = record;

      // dateとhashが必要な場合に処理を進める
      if (!date || !hash) {
        console.warn(`シート "${sheetName}" に不完全なデータがあります:`, record);
        continue;
      }

      // ファイル名と保存パスを生成
      const fileName = `${sheetName}-${date}-${hash}.png`;
      const sheetFolder = path.join(backupFolder, sheetName);
      const savePath = path.join(sheetFolder, fileName);

      // 保存先フォルダが存在しない場合は作成
      await fsp.mkdir(sheetFolder, { recursive: true });

      // ファイルが存在するか確認
      if (await fileExists(savePath)) {
        console.log(`ファイルが既に存在します: ${savePath}`);
        continue; // 次のレコードにスキップ
      }

      // URLを生成してダウンロード
      const imageUrl = `https://pbs.twimg.com/media/${hash}?format=png&name=4096x4096`;
      console.log(`ダウンロード中: ${imageUrl} → ${savePath}`);
      try {
        await downloadImage(imageUrl, savePath);
        console.log(`保存成功: ${savePath}`);
      } catch (error) {
        console.error(`ダウンロードエラー: ${imageUrl}`, error);
      }
    }
  }
}

// スクリプトの実行
(async () => {
  try {
    await processExcel(excelFile);
    console.log('処理が完了しました。');
  } catch (error) {
    console.error('エラーが発生しました:', error);
  }
})();
