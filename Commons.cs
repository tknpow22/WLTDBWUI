using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WLTDBWUI
{
    static internal class Commons
    {
        /// <summary>
        /// ステータス表示用のテキストボックス
        /// </summary>
        private static TextBox textBoxStatus;

        /// <summary>
        /// ステータス表示用のテキストボックスを設定する
        /// </summary>
        /// <param name="tBoxStatus"></param>
        public static void SetStatusTextBox(TextBox tBoxStatus)
        {
            textBoxStatus = tBoxStatus;
        }

        /// <summary>
        /// プログラムからの出力用
        /// </summary>
        /// <param name="format"></param>
        /// <param name="arg"></param>
        public static void WriteLine(string format, params object[] arg)
        {
            string text = string.Format(format, arg);
            text = text.Replace("\r", "");
            text = text.Replace("\n", "");

            if (textBoxStatus == null) {
                return;
            }                
                    
            if (textBoxStatus.InvokeRequired) {
                Action invokeFunction = delegate { WriteLine(text); };
                textBoxStatus.Invoke(invokeFunction);
            } else {
                textBoxStatus.Text = text;
            }
        }
    }
}
