using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using Stats = DerPapierjosef.JosefStatistics.Stats;


namespace DerPapierjosef
{
    public partial class JosefRibbon
    {
        JosefModel model;
        OpenNLP nlp;

        private void JosefRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Josef_Click(object sender, RibbonControlEventArgs e)
        {
            Action exec = loadNLP,
                    exec2 = updateModel;
            ProgressForm pf = new ProgressForm();
            BackgroundWorker b = new BackgroundWorker(),
                c = new BackgroundWorker();
            b.DoWork += (object s, DoWorkEventArgs ea) =>
            {
                exec.Invoke();
            };
            c.DoWork += (object s2, DoWorkEventArgs ea2) =>
            {
                exec2.Invoke();
            };
            b.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs ea) =>
            {
                pf.label1.Text="Text wird ausgewertet... (Schritt 2 von 2)";
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

        private void showAuswertungen()
        {
            Globals.ThisAddIn.TaskPane.Visible = true;
            Globals.ThisAddIn.TaskPane.Width = 344;
            JosefStatistics stats = new JosefStatistics(model);
            Auswertung aw = new Auswertung();
            aw.Show();
            setLabels(aw, stats);
            fillHistogram(aw);
            fillGridViews(aw);
            Globals.ThisAddIn.MyJosefPane.setNodes(buildTreeNodes(stats));
            Globals.ThisAddIn.MyJosefPane.treeView1.NodeMouseClick +=
            (object sender, System.Windows.Forms.TreeNodeMouseClickEventArgs e) =>
            {
                SentenceTreeNode stn = (SentenceTreeNode)e.Node;
                model.Document.Range(stn.begin, stn.end).Select();
            };
        }
                
        void fillHistogram(Auswertung aw)
        {
            int[] lengths = model.Sentences.Select(s => s.Words.Count).ToArray();
            for (int i = lengths.Min(); i < lengths.Max() + 1; i++)
            {
                aw.chart1.Series[0].Points.AddXY(i, lengths.Where(l => l == i).Count());
            }
        }

        void setLabels(Auswertung aw,JosefStatistics stats)
        {
            aw.label3.Text = "" + stats.BasicStatistics.wordCount;
            aw.label12.Text = "" + stats.BasicStatistics.sentenceCount;
            aw.label4.Text = "" + stats.BasicStatistics.hardSentenceCount;
            aw.label13.Text = "" + stats.BasicStatistics.paragraphCount;
            aw.label7.Text = "" + stats.BasicStatistics.dickesSteiwer + "%";
            aw.label16.Text = "" + stats.BasicStatistics.floskelMean + "%";
            aw.label8.Text = "" + stats.BasicStatistics.nominalMean + "%";
            aw.label9.Text = "" + stats.BasicStatistics.uniqueWordCount;
            aw.label28.Text = "" + stats.korrelatSaetze.Count;
            aw.label29.Text = "" + stats.passivSaetze.Count;
            aw.label30.Text = "" + stats.vergangenheitsSätze.Count;
            aw.label31.Text = "" + stats.unpersoenlicheSaetze.Count;
            aw.label32.Text = "" + stats.komplexeSaetze.Count;
            aw.label33.Text = "" + stats.langeSaetze.Count;
            aw.label34.Text = "" + stats.phrasenSaetze.Count;
            aw.label35.Text = "" + stats.fuellwortSaetze.Count;
            aw.label36.Text = "" + stats.nominalSaetze.Count;
        }

        void fillGridViews(Auswertung aw1)
        {
            foreach (KeyValuePair<string, int> kv in model.FrequentTags.OrderByDescending(s => s.Value).Take(100))
            {
                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
                row.CreateCells(aw1.dataGridView1);
                row.Cells[0].Value = kv.Key;
                row.Cells[1].Value = kv.Value;
                aw1.dataGridView1.Rows.Add(row);
            }

            foreach (KeyValuePair<string, int> v in model.UniqueWords.OrderByDescending(s => s.Value).Take(100))
            {
                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
                row.CreateCells(aw1.dataGridView2);
                row.Cells[0].Value = v.Key;
                row.Cells[1].Value = v.Value;
                aw1.dataGridView2.Rows.Add(row);
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
                SentenceTreeNode tn = new SentenceTreeNode();
                tn.Text = satz.Text.Substring(0, Math.Min(20, satz.Text.Length)) + "...";
                tn.begin = satz.Start;
                tn.end = satz.End;
                fuellNodes.Add(tn);
            }
            foreach (Word.Range satz in statistics.langeSaetze)
            {
                SentenceTreeNode tn = new SentenceTreeNode();
                tn.Text = satz.Text.Substring(0, Math.Min(20, satz.Text.Length)) + "...";
                tn.begin = satz.Start;
                tn.end = satz.End;
                langNodes.Add(tn);
            }
            foreach (Word.Range satz in statistics.komplexeSaetze)
            {
                SentenceTreeNode tn = new SentenceTreeNode();
                tn.Text = satz.Text.Substring(0, Math.Min(20, satz.Text.Length)) + "...";
                tn.begin = satz.Start;
                tn.end = satz.End;
                komplexNodes.Add(tn);
            }
            foreach (Word.Range satz in statistics.passivSaetze)
            {
                SentenceTreeNode tn = new SentenceTreeNode();
                tn.Text = satz.Text.Substring(0, Math.Min(20, satz.Text.Length)) + "...";
                tn.begin = satz.Start;
                tn.end = satz.End;
                passivNodes.Add(tn);
            }
            int sc = statistics.BasicStatistics.sentenceCount;
            return new SentenceTreeNode[] { new SentenceTreeNode("Füllwörter-Sätze ("+fuellNodes.Count+" von "+sc+")", fuellNodes.ToArray()),
                                            new SentenceTreeNode("Lange Sätze ("+langNodes.Count+" von "+sc+")", langNodes.ToArray()),
                                            new SentenceTreeNode("Komplexe Sätze ("+komplexNodes.Count+" von "+sc+")", komplexNodes.ToArray()),
                                            new SentenceTreeNode("Passiv-Sätze ("+passivNodes.Count+" von "+sc+")", passivNodes.ToArray())};
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

        private void updateModel()
        {
            model = new JosefModel(Globals.ThisAddIn.Application.ActiveDocument, nlp);
        }
    }
}
