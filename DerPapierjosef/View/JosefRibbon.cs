using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using Stats = DerPapierjosef.JosefStatistics.Stats;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms;



namespace DerPapierjosef
{
    public partial class JosefRibbon
    {
        JosefModel model;
        OpenNLP nlp;
        BackgroundWorker c;
        ProgressForm pf;

        private void JosefRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Josef_Click(object sender, RibbonControlEventArgs e)
        {
            Action exec = loadNLP,
                    exec2 = updateModel;
            pf = new ProgressForm();
            BackgroundWorker b = new BackgroundWorker();
            c = new BackgroundWorker();
            b.DoWork += (object s, DoWorkEventArgs ea) =>
            {
                exec.Invoke();
            };
            c.DoWork += (object s2, DoWorkEventArgs ea2) =>
            {
                exec2.Invoke();
            };
            c.WorkerReportsProgress = true;
            c.ProgressChanged += c_ProgressChanged;
            b.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs ea) =>
            {
                pf.label1.Text="Text wird ausgewertet... (Schritt 2 von 2)";
                pf.progressBar1.MarqueeAnimationSpeed = 0;
                pf.progressBar1.Style = ProgressBarStyle.Continuous;
                c.RunWorkerAsync();
            };
            c.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs ea) =>
            {
                if (pf != null && pf.Visible) pf.Close();
                showAuswertungen();
            };
            pf.Show();
            pf.progressBar1.MarqueeAnimationSpeed = 30;
            b.RunWorkerAsync();
            
        }

        void c_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pf.progressBar1.Value = e.ProgressPercentage;
        }

        private void showAuswertungen()
        {
            JosefPane MyJosefPane = Globals.ThisAddIn.MyJosefPane;
            Globals.ThisAddIn.TaskPane.Visible = true;
            Globals.ThisAddIn.TaskPane.Width = 344;
            JosefStatistics stats = new JosefStatistics(model);
            setLabels(stats,MyJosefPane.label7,MyJosefPane.label8,MyJosefPane.label9,MyJosefPane.label10);
            
            fillHistogram(MyJosefPane.chart1);
            fillGridViews(MyJosefPane.dataGridView2);
            MyJosefPane.setNodes(buildTreeNodes(stats));
            MyJosefPane.treeView1.NodeMouseClick +=
            (object sender, System.Windows.Forms.TreeNodeMouseClickEventArgs e) =>
            {
                SentenceTreeNode stn = (SentenceTreeNode)e.Node;
                model.Document.Range(stn.begin, stn.end).Select();
            };
        }
                
        void fillHistogram(Chart chart1)
        {

            int[] lengths = model.Sentences.Select(s => s.Words.Count).ToArray();
            for (int i = lengths.Min(); i < lengths.Max() + 1; i++)
            {
                chart1.Series[0].Points.AddXY(i, lengths.Where(l => l == i).Count());
            }
        }

        void setLabels(JosefStatistics stats, Label lbl7, Label lbl8, Label lbl9, Label lbl10)
        {
            lbl7.Text = "" + stats.BasicStatistics.paragraphCount;
            lbl9.Text = "" + stats.BasicStatistics.wordCount;
            lbl8.Text = "" + stats.BasicStatistics.sentenceCount;
            lbl10.Text = "" + stats.BasicStatistics.uniqueWordCount;
        }

        void fillGridViews(System.Windows.Forms.DataGridView dataGridView2)
        {
            float wcount = model.Words.Count;
            foreach (KeyValuePair<string, int> v in model.UniqueWords.OrderByDescending(s => s.Value).Take(100))
            {
                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
                row.CreateCells(dataGridView2);
                row.Cells[0].Value = v.Key;
                row.Cells[1].Value = v.Value;
                row.Cells[2].Value = Math.Round(100.0f*v.Value / wcount,2)+"%";
                dataGridView2.Rows.Add(row);
            }
        }

        public class SentenceTreeNode : System.Windows.Forms.TreeNode
        {
            public int begin, end;
            public SentenceTreeNode(string name, SentenceTreeNode[] nodes) : base(name,nodes)
            {
            }

            public SentenceTreeNode()
                : base()
            {

            }
        }

        SentenceTreeNode[] buildTreeNodes(JosefStatistics statistics)
        {
            List<SentenceTreeNode> fuellNodes = new List<SentenceTreeNode>();
            List<SentenceTreeNode> langNodes = new List<SentenceTreeNode>();
            List<SentenceTreeNode> komplexNodes = new List<SentenceTreeNode>();
            List<SentenceTreeNode> passivNodes = new List<SentenceTreeNode>();
            foreach (Word.Range satz in statistics.fuellwortSaetze)
            {
                fuellNodes = addNodeToList(satz, fuellNodes, true);
            }
            foreach (Word.Range satz in statistics.langeSaetze)
            {
                langNodes=addNodeToList(satz, langNodes, true);
            }
            foreach (Word.Range satz in statistics.komplexeSaetze)
            {
                komplexNodes = addNodeToList(satz, komplexNodes, false);
            }
            foreach (Word.Range satz in statistics.passivSaetze)
            {
                passivNodes=addNodeToList(satz, passivNodes, true);
            }
            return buildParentNodes(statistics.BasicStatistics.sentenceCount,fuellNodes,langNodes,komplexNodes,passivNodes);
        }

        private SentenceTreeNode[] buildParentNodes(int sc,List<SentenceTreeNode> fuellNodes,
                                                           List<SentenceTreeNode> langNodes,
                                                           List<SentenceTreeNode> komplexNodes,
                                                           List<SentenceTreeNode> passivNodes)
        {
            string end=" von " + sc + ")";
            SentenceTreeNode fuellnode = new SentenceTreeNode("Füllwörter-Sätze (" + fuellNodes.Count + end, fuellNodes.ToArray()),
                             langnode = new SentenceTreeNode("Lange Sätze (" + langNodes.Count + end, langNodes.ToArray()),
                             komplexnode = new SentenceTreeNode("Komplexe Sätze (" + komplexNodes.Count + end, komplexNodes.ToArray()),
                             passivnode = new SentenceTreeNode("Passiv-Sätze (" + passivNodes.Count + end, passivNodes.ToArray());
            fuellnode.BackColor = ((float)fuellNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)fuellNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            langnode.BackColor = ((float)langNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)langNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            komplexnode.BackColor = ((float)komplexNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)komplexNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            passivnode.BackColor = ((float)passivNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)passivNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            return new SentenceTreeNode[] { fuellnode,
                                            langnode,
                                            komplexnode,
                                            passivnode};
        }

        private void loadNLP()
        {
            Globals.ThisAddIn.Application.Selection.WholeStory();
            string path = "C:\\Users\\Dominik\\Documents\\visual studio 2013\\Projects\\DerPapierjosef\\DerPapierjosef\\";
            if (nlp == null) 
            {
                nlp = new OpenNLP(path);
            }
        }

        private List<SentenceTreeNode> addNodeToList(Word.Range satz, List<SentenceTreeNode> nodeList,bool showNoWords)
        {
            const int SENTENCE_LENGTH = 35;
            SentenceTreeNode tn = new SentenceTreeNode();
            tn.Text = satz.Text.Substring(0, Math.Min(SENTENCE_LENGTH, satz.Text.Length)) + "..."
                +(showNoWords?(" ("+satz.Words.Count+")"):"");
            tn.begin = satz.Start;
            tn.end = satz.End;
            nodeList.Add(tn);
            return nodeList;
        }

        private void updateModel()
        {
            model = new JosefModel(Globals.ThisAddIn.Application.ActiveDocument, nlp,c);
        }
    }
}
