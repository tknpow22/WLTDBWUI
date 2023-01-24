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
        /// コンソール用のテキストボックス
        /// </summary>
        private static TextBox textBoxConsole;

        /// <summary>
        /// コンソール用のテキストボックスを設定する
        /// </summary>
        /// <param name="tBoxConsole"></param>
        public static void SetConsoleTextBox(TextBox tBoxConsole)
        {
            textBoxConsole = tBoxConsole;
        }

        /// <summary>
        /// プログラムからの出力を行う
        /// </summary>
        /// <param name="format"></param>
        /// <param name="arg"></param>
        public static void WriteLine(string format, params object[] arg)
        {
            if (textBoxConsole == null) {
                return;
            }                
                    
            if (textBoxConsole.InvokeRequired) {
                Action invokeFunction = delegate { WriteLine(format, arg); };
                textBoxConsole.Invoke(invokeFunction);
            } else {
                textBoxConsole.AppendText(string.Format(format + "\r\n", arg));
            }
        }
    }
}
