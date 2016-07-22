using Microsoft.Scripting.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Windows.Forms;

namespace WordAnalysis
{
    public partial class LineForm : Form
    {
        public LineForm()
        {
            InitializeComponent();
        }
        public LineForm(string text)
        {
            InitializeComponent();
            input = text;
        }
        private List<string> keywords = new List<string>();
        private string input = @"
3、用餐	全程含2早餐4正餐；    正餐30元/人/餐 *4*30                          3600    
4、住宿	千岛湖秀水度假酒店   420元/间/晚（含早）*2*15                        12600     
4、导游	中文导游服务；       300元/天/车*3                                     900
5、保险	旅行社责任险,旅游人身意外伤害险    10元/人/天*3*30                     900
6、其他	综合服务费； 5元/人/天*3*30                                            450
";
        private void LineForm_Load(object sender, EventArgs e)
        {
            Text = "请选择出资源名称";
            var chsPattern = @"[\u4e00-\u9fa5]+";
            var numberPattern = @"([\,\.\+\-\*\/\^\=]*[0-9]+)+";
            var pin = new Point(20, 30);
            var offset = 10;
            var fontSize = 14.25F;
            var font = new Font("微软雅黑", fontSize, FontStyle.Regular, GraphicsUnit.Pixel, 134);
            foreach (Button button in Controls.OfType<Button>().Where(b => b.FlatStyle == FlatStyle.Flat))
            {
                Controls.Remove(button);
            }
            Controls.AddRange(
                Matches(input, chsPattern, numberPattern).Select(
                    (s, i) =>
                    {
                        var size = new Size(30 * s.Length, 45);
                        var location = new Point(pin.X, pin.Y);
                        if (pin.X+size.Width+offset>Width)
                        {
                            pin.X = 20;
                            pin.Y += 60;
                            location = new Point(pin.X, pin.Y);
                        }
                        pin.X += size.Width + offset;
                        var button = new Button
                                         {
                                             Font = font,
                                             Text = s,
                                             Size = size,
                                             //Location = new Point(pin.X, pin.Y + i * (offset + size.Height)),
                                             Location = location,
                                             FlatStyle = FlatStyle.Flat,
                                         };
                        button.Click += new EventHandler(
                            (o, ev) =>
                                {
                                    Button current = o as Button;
                                    if (current.ForeColor == Color.Blue)
                                    {
                                        current.ForeColor = Color.Black;
                                    }
                                    else
                                    {
                                        current.ForeColor = Color.Blue;
                                    }
                                });
                        return button;
                    }).ToArray());
        }

        private static IEnumerable<string> Matches(string s, params string[] patterns)
        {
            int x = 0;
            List<int> list =
                patterns.Select(p => s.Matches(p))
                    .SelectMany(p => p)
                    .Select(t => new[] { x = s.IndexOf(t, x), x += t.Length })
                    .SelectMany(p => p)
                    .Union(new[] { s.Length })
                    .Where(p => p >= 0)
                    .ToList()
                    .HashSet()
                    .OrderBy(p => p)
                    .ToList();
            for (int i = 0; i < list.Count - 1; i ++)
            {
                yield return s.Substring(list[i], list[i + 1] - list[i]).Trim();
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            int idx = 0;
            MessageBox.Show(
                Controls.OfType<Button>()
                    .Where(b => b.ForeColor == Color.Blue)
                    .Select(b => b.Text)
                    .OrderBy(k => idx = input.IndexOf(k, idx))
                    .JoinStrings(" "));
        }

        private void LineForm_Resize(object sender, EventArgs e)
        {
            OnLoad(e);
        }
    }
}
