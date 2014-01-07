using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DerPapierjosef
{
    public partial class JosefPane : UserControl
    {
        

        public JosefPane()
        {
            InitializeComponent();
        }

        private void JosefPane_Load(object sender, EventArgs e)
        {
            TreeNode[] tn1={new TreeNode("Dieser Satz hat Flossen"),new TreeNode("Hoher Floskel-Faktor"),new TreeNode("Lass die Floskeln tanzen")};
            TreeNode tn=new TreeNode("Floskeln",tn1);
            tn.BackColor = Color.Yellow;
            treeView1.Nodes.Add(tn);
            TreeNode tn2 = new TreeNode("Nominalstil");
            tn2.BackColor = Color.LawnGreen;
            treeView1.Nodes.Add(tn2);
        }

        public void setNodes(TreeNode[] tn)
        {
            treeView1.Nodes.Clear();
            treeView1.Nodes.AddRange(tn);
        }
    }
}
