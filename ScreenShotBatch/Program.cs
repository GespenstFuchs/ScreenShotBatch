using IronSoftware.Drawing;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ScreenShotBatch
{
    /// <summary>
    /// スクリーンショットバッチ
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// メイン処理
        /// </summary>
        /// <param name="args">
        /// 0：キャプチャ開始位置：Ｘ座標
        /// 1：キャプチャ開始位置：Ｙ座標
        /// 2：キャプチャ幅
        /// 3：キャプチャ高さ
        /// 4：保存フォルダパス
        /// 5：連番桁数（省略可能（デフォルト：３桁））
        /// 
        /// ※：パラメータは、半角スペースで区切る。
        /// ※：保存フォルダパスや、ベースファイル名にスペースが含まれる場合、ダブルコーテーションで囲む。、
        /// </param>
        static void Main(string[] args)
        {
            // テスト用コード
            // （使用しているKindleに合わせたサイズ）
            //args =
            //[
            //    "48",  // キャプチャ開始位置：Ｘ座標
            //    "100",  // キャプチャ開始位置：Ｙ座標
            //    "1870", // キャプチャ幅
            //    "888", // キャプチャ高さ
            //    @"C:\Users\fugyu\OneDrive\画像", // 保存フォルダパス
            //    "3",  // 連番桁数（省略可能）
            //];
            // テスト用コード

            // パラメータ数チェックを行う。
            if (5 > args.Length)
            {
                return;
            }

            // 保存フォルダパスの有無を判定する。
            if (!Directory.Exists(args[4]))
            {
                return;
            }

            // 連番桁数を設定する。
            int length = 3;
            if (6 <= args.Length)
            {
                // 連番桁数を取得する。
                if (!int.TryParse(args[5], out length) || length < 0)
                {
                    return;
                }
            }

            string savaFolderPath = @"C:\Users\fugyu\OneDrive\画像";
            string extension = ".png";

            // キャプチャする範囲を指定する。
            System.Drawing.Rectangle captureArea = new(
                int.Parse(args[0]),
                int.Parse(args[1]),
                int.Parse(args[2]),
                int.Parse(args[3]));

            // 正規表現を生成する。
            //（数値（ファイル名）＋拡張子）
            Regex regex = new(string.Concat("^", @"\d{", length, "}", Regex.Escape(extension), @"$"));

            // 該当ファイルを取得する。
            var filesList = Directory.GetFiles(savaFolderPath)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name) && regex.IsMatch(name))
                .ToList();

            int maxNumber = 0;

            // 該当ファイルの有無を判定する。
            if (filesList.Any())
            {
                // 正規表現を生成する。
                //（数値（ファイル名））
                regex = new(string.Concat("^", @"\d{", length, "}"));

                // 最大番号を抽出する。
                maxNumber = filesList
                    .Select(name => int.Parse(regex.Match(name!).Groups[0].Value))
                    .DefaultIfEmpty(0)
                    .Max();

                // 最大番号＋１の桁数を判定する。
                if ((maxNumber + 1).ToString().Length <= length)
                {
                    // 桁数を超えない場合、最大番号＋１を設定する。
                    maxNumber = maxNumber + 1;
                }
            }

            // ファイル名を生成する。
            // （ベースファイル名は、半角数値を使用されると、処理が複雑になるので、使用しない。）
            string nextFileName = $"{maxNumber.ToString(string.Concat("D", length))}{extension}";

            using Bitmap bmp = new(captureArea.Width, captureArea.Height);
            using var g = Graphics.FromImage(bmp);

            // パラメータについて、
            // 最初の２個：キャプチャ元のスクリーン上のＸ・Ｙ座標（ハードコピー取得位置）
            // 次の２個：キャプチャ先のスクリーン上のＸ・Ｙ座標（描画位置）
            // 最後のパラメータ：キャプチャする範囲のサイズ（幅・高さ）
            //g.CopyFromScreen(0, 0, 0, 0, captureArea.Size);
            g.CopyFromScreen(
                int.Parse(args[0]),
                int.Parse(args[1]),
                0,
                0,
                captureArea.Size);

            #region .NETのみ

            //// スクリーンショットを保存する。
            //bmp.Save(Path.Combine(savaFolderPath, nextFileName), ImageFormat.Png);

            #endregion

            // BitmapをAnyBitmapに変換する。
            AnyBitmap anyBmp = AnyBitmap.FromBitmap(bmp);

            // スクリーンショットを保存する。
            anyBmp.ExportFile(Path.Combine(savaFolderPath, nextFileName), AnyBitmap.ImageFormat.Png, 100);
        }
    }
}
