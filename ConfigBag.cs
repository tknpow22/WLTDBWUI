using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WLTDBWUI
{
    /// <summary>
    /// 設定処理
    /// </summary>
    internal class ConfigBag
    {
        #region Windows API

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern int GetPrivateProfileString(
               string lpApplicationName,
               string lpKeyName,
               string lpDefault,
               StringBuilder lpReturnedstring,
               int nSize,
               string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern int WritePrivateProfileString(
            string lpApplicationName,
            string lpKeyName,
            string lpString,
            string lpFileName);

        #endregion

        #region メンバ

        /// <summary>
        /// 実行モジュールのあるディレクトリ
        /// </summary>
        private string appDirectory = "";

        /// <summary>
        /// 設定ファイルのパス
        /// </summary>
        private string configFilepath = "";

        /// <summary>
        /// セクション名 - エンコーディングの警告用
        /// </summary>
        private const string SECTION_NAME_ENCODING_WARNING = "EncodingWarning";

        /// <summary>
        /// セクション名 - 設定
        /// </summary>
        private const string SECTION_NAME = "Config";

        #region キー名 - エンコーディングの警告用

        private const string KEY_ENCODING_WARN_WARNING = "EncodingWarning";

        #endregion


        #region キー名 - 設定

        /// <summary>
        /// データベースディレクトリ
        /// </summary>
        private const string KEY_DB_DIRECTORY = "DbDirectory";

        /// <summary>
        /// データベースファイル名
        /// </summary>
        private const string KEY_DB_FILE_NAME = "DbFilename";

        /// <summary>
        /// ログディレクトリ
        /// </summary>
        private const string KEY_LOG_DIRECTORY = "LogDirectory";

        /// <summary>
        /// バックアップディレクトリ
        /// </summary>
        private const string KEY_LOG_BACKUP_DIRECTORY = "BackupDirectory";

        /// <summary>
        /// WLID の一覧
        /// </summary>
        private const string KEY_WLIDS_TSV = "WlIdsTsv";

        /// <summary>
        /// WLID のエイリアス名の一覧
        /// </summary>
        private const string KEY_WLID_ALIASES_TSV = "WlIdAliasesTsv";

        /// <summary>
        /// 出力する CSV の日付の日の範囲(差)
        /// </summary>
        private const string KEY_OUTPUT_CSV_DATE_DAY_RANGE = "CsvDateDayRange";

        /// <summary>
        /// CSV を出力するディレクトリ名
        /// </summary>
        private const string KEY_CSV_DIRECTORY = "CsvDirectory";

        /// <summary>
        /// 出力する CSV のファイル名
        /// </summary>
        private const string KEY_CSV_FILE_NAME = "CsvFilename";

        /// <summary>
        /// CSV 作成後関連付け起動する
        /// </summary>
        private const string KEY_EXEC_CSV = "ExecCsv";

        #endregion

        #endregion

        #region プロパティ

        /// <summary>
        /// データベースディレクトリ
        /// </summary>
        public string DbDirectory
        {
            get; private set;
        }

        /// <summary>
        /// データベースファイル名
        /// </summary>
        public string DbFilename
        {
            get; private set;
        }

        /// <summary>
        /// データベースファイルパス
        /// </summary>
        public string DbFilepath
        {
            get
            {
                return Path.Combine(this.DbDirectory, this.DbFilename);
            }
        }

        /// <summary>
        /// ログディレクトリ
        /// </summary>
        public string LogDirectory
        {
            get; private set;
        }

        /// <summary>
        /// バックアップディレクトリ
        /// </summary>
        public string LogBackupDirectory
        {
            get; private set;
        }

        /// <summary>
        /// WLIDエイリアス名連想配列
        /// </summary>
        public Dictionary<string, string> WlIdAliases
        {
            get; set;
        }

        /// <summary>
        /// 設定で保存する際の WLID の順番に格納する
        /// </summary>
        public List<string> WlIdByOrder
        {
            get; set;
        }

        /// <summary>
        /// 出力する CSV の日付の日の範囲(差)
        /// </summary>
        public int OutputCsvDateDayRange
        {
            get; set;
        }

        /// <summary>
        /// CSV を出力するディレクトリ名
        /// </summary>
        public string CsvDirectory
        {
            get; private set;
        }

        /// <summary>
        /// 出力する CSV のファイル名
        /// </summary>
        public string CsvFilename
        {
            get; private set; 
        }

        /// <summary>
        /// CSV 作成後関連付け起動する
        /// </summary>
        public bool ExecCSV
        {
            get; set;
        }

        #endregion

        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="appDirectory">実行モジュールのあるディレクトリ</param>
        public void Initialize(string appDirectory)
        {
            this.appDirectory = appDirectory;
            this.configFilepath = Path.Combine(appDirectory, "config.ini");

            this.Load();
        }

        /// <summary>
        /// 設定を読み込む
        /// </summary>
        private void Load()
        {
            {
                {
                    bool valid = true;

                    string dbDirectoryText = this.GetString(SECTION_NAME, KEY_DB_DIRECTORY);
                    if (string.IsNullOrEmpty(dbDirectoryText)) {
                        valid = false;
                    } else if (!Directory.Exists(dbDirectoryText)) {
                        // フルパスが指定されていなかった時を想定
                        if (0 <= dbDirectoryText.IndexOfAny(Path.GetInvalidPathChars())) {
                            valid = false;
                        }
                    }
                    if (!valid) {
                        dbDirectoryText = Path.Combine(this.appDirectory, "Db");
                    }

                    this.DbDirectory = dbDirectoryText;
                }
            }

            {
                string dbFilenameText = this.GetString(SECTION_NAME, KEY_DB_FILE_NAME);
                if (string.IsNullOrEmpty(dbFilenameText)) {
                    dbFilenameText = "WLTDB.db3";
                }
                this.DbFilename = dbFilenameText;
            }

            {
                bool valid = true;

                string logDirectoryText = this.GetString(SECTION_NAME, KEY_LOG_DIRECTORY);
                if (string.IsNullOrEmpty(logDirectoryText)) {
                    valid = false;
                } else if (!Directory.Exists(logDirectoryText)) {
                    if (0 <= logDirectoryText.IndexOfAny(Path.GetInvalidPathChars())) {
                        valid = false;
                    }
                }
                if (!valid) {
                    logDirectoryText = Path.Combine(this.appDirectory, "Log");
                }

                this.LogDirectory = logDirectoryText;
            }

            {
                bool valid = true;

                string logBackupDirectoryText = this.GetString(SECTION_NAME, KEY_LOG_BACKUP_DIRECTORY);
                if (string.IsNullOrEmpty(logBackupDirectoryText)) {
                    valid = false;
                } else if (!Directory.Exists(logBackupDirectoryText)) {
                    if (0 <= logBackupDirectoryText.IndexOfAny(Path.GetInvalidPathChars())) {
                        valid = false;
                    }
                }
                if (!valid) {
                    logBackupDirectoryText = Path.Combine(this.LogDirectory, "Backup");
                }

                this.LogBackupDirectory = logBackupDirectoryText;
            }

            {
                string wlIdsTSVText = this.GetString(SECTION_NAME, KEY_WLIDS_TSV);
                string wlIdAliasesTSVText = this.GetString(SECTION_NAME, KEY_WLID_ALIASES_TSV);
                
                string[] wlIdList = wlIdsTSVText.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                string[] wlIdAliasList = wlIdAliasesTSVText.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);

                Dictionary<string, string> wlIdAliases = new Dictionary<string, string>();
                List<string> wlIdByOrder = new List<string>();

                for (int i = 0; i < wlIdList.Count(); ++i) {
                    if (wlIdAliasList.Count() <= i) {
                        break;
                    }

                    wlIdAliases.Add(wlIdList[i], wlIdAliasList[i]);
                    wlIdByOrder.Add(wlIdList[i]);
                }
                this.WlIdAliases = wlIdAliases;
                this.WlIdByOrder = wlIdByOrder;
            }

            {
                this.OutputCsvDateDayRange = this.GetInt(SECTION_NAME, KEY_OUTPUT_CSV_DATE_DAY_RANGE, 3);
            }

            {
                bool valid = true;

                string csvDirectoryText = this.GetString(SECTION_NAME, KEY_CSV_DIRECTORY);
                if (string.IsNullOrEmpty(csvDirectoryText)) {
                    valid = false;
                } else if (!Directory.Exists(csvDirectoryText)) {
                    if (0 <= csvDirectoryText.IndexOfAny(Path.GetInvalidPathChars())) {
                        valid = false;
                    } else {
                        csvDirectoryText = Path.Combine(this.appDirectory, csvDirectoryText);
                    }
                }
                if (!valid) {
                    csvDirectoryText = this.appDirectory;
                }

                this.CsvDirectory = csvDirectoryText;
            }

            {
                string csvFileNameText = this.GetString(SECTION_NAME, KEY_CSV_FILE_NAME);
                if (string.IsNullOrEmpty(csvFileNameText)) {
                    csvFileNameText = "Temperature.csv";
                }
                this.CsvFilename = csvFileNameText;
            }

            {
                //ExecCSV
                string execCSVText = this.GetString(SECTION_NAME, KEY_EXEC_CSV);
                if (string.IsNullOrEmpty(execCSVText)) {
                    execCSVText = "False";
                }
                this.ExecCSV = execCSVText.ToUpper() == "TRUE";
            }
        }

        /// <summary>
        /// 設定を保存する
        /// </summary>
        public void Save()
        {
            {
                this.WriteString(SECTION_NAME_ENCODING_WARNING, KEY_ENCODING_WARN_WARNING, "このファイルは必ず Shift JIS で保存してください。");
            }

            {
                List<string> wlIdAliasList = new List<string>();
                foreach (string wlId in this.WlIdByOrder) {
                    string alias = this.WlIdAliases[wlId];
                    wlIdAliasList.Add(alias);
                }

                string wlIdsTSVText = string.Join("\t", this.WlIdByOrder);
                string wlIdAliasesTSVText = string.Join("\t", wlIdAliasList);

                this.WriteString(SECTION_NAME, KEY_WLIDS_TSV, wlIdsTSVText);
                this.WriteString(SECTION_NAME, KEY_WLID_ALIASES_TSV, wlIdAliasesTSVText);
            }

            this.WriteInt(SECTION_NAME, KEY_OUTPUT_CSV_DATE_DAY_RANGE, this.OutputCsvDateDayRange);

            this.WriteString(SECTION_NAME, KEY_EXEC_CSV, this.ExecCSV ? "True" : "False");
        }

        /// <summary>
        /// 設定ファイルから文字列を読み込む
        /// </summary>
        /// <param name="lpApplicationName">セクション名</param>
        /// <param name="lpKeyName">キー名</param>
        /// <param name="lpDefault">デフォルト値</param>
        /// <returns>読み込んだ値</returns>
        private string GetString(string lpApplicationName, string lpKeyName, string lpDefault = "")
        {
            StringBuilder buffer = new StringBuilder(2560);
            GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, buffer, buffer.Capacity, this.configFilepath);
            return buffer.ToString();
        }

        /// <summary>
        /// 設定ファイルから数値を取得する
        /// </summary>
        /// <param name="lpApplicationName">セクション名</param>
        /// <param name="lpKeyName">キー名</param>
        /// <param name="nDefault">デフォルト値</param>
        /// <returns></returns>
        private int GetInt(string lpApplicationName, string lpKeyName, int nDefault = 0)
        {
            string str = this.GetString(lpApplicationName, lpKeyName, "");
            int result;

            if (int.TryParse(str, out result)) {
                return result;
            }

            return nDefault;
        }

        /// <summary>
        /// 設定ファイルに文字列を書き込む
        /// </summary>
        /// <param name="lpApplicationName">セクション名</param>
        /// <param name="lpKeyName">キー名</param>
        /// <param name="lpString">デフォルト値</param>
        private void WriteString(string lpApplicationName, string lpKeyName, string lpString)
        {
            WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, this.configFilepath);
        }

        /// <summary>
        /// 設定ファイルに数値を文字列として書き込む
        /// </summary>
        /// <param name="lpApplicationName">セクション名</param>
        /// <param name="lpKeyName">キー名</param>
        /// <param name="nInt">デフォルト値</param>
        private void WriteInt(string lpApplicationName, string lpKeyName, int nInt)
        {
            string sInt = string.Format("{0}", nInt);
            this.WriteString(lpApplicationName, lpKeyName, sInt);
        }

    }
}
