﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using Stats = DerPapierjosef.JosefStatistics.Stats;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms;
using DerPapierjosef.View;




namespace DerPapierjosef
{
    public partial class JosefRibbon
    {
        JosefModel model;
        OpenNLP nlp;
        BackgroundWorker c;
        ProgressForm pf;
        JosefStatistics stats;
        OptionsDialog dlg;
        TimerClass tc;

        private void JosefRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void Josef_Click(object sender, RibbonControlEventArgs e)
        {
            pf = new ProgressForm();
            BackgroundWorker b = onlpLoader();
            c = textAnalyzer();
            pf.label5.Text = "";
            pf.Show();
            tc = new TimerClass();
            tc.StartClock(ref pf.label3);
            pf.progressBar1.MarqueeAnimationSpeed = 30;
            b.RunWorkerAsync();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            fillGridViewFreq(Globals.ThisAddIn.MyJosefPane.dataGridView2,
                          Globals.ThisAddIn.MyJosefPane.checkBox1.Checked);
        }

        private BackgroundWorker onlpLoader()
        {
            BackgroundWorker b = new BackgroundWorker();

            b.DoWork += (object s, DoWorkEventArgs ea) =>
            {
                Action exec = loadNLP;
                exec.Invoke();
            };
            b.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs ea) =>
            {
                pf.label1.Text = "Text wird ausgewertet...";
                pf.label7.Text = "2 von 2";
                pf.progressBar1.Style = ProgressBarStyle.Continuous;
                c.RunWorkerAsync();
            };
            return b;
        }

        private BackgroundWorker textAnalyzer()
        {
            BackgroundWorker cc=new BackgroundWorker();
            cc.WorkerReportsProgress = true;
            cc.ProgressChanged += (object sender1, ProgressChangedEventArgs e1) =>
            {
                pf.progressBar1.Value = (int)100*e1.ProgressPercentage / Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count;
                pf.label5.Text = e1.ProgressPercentage + " von " + Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count;
            };
            cc.DoWork += (object s2, DoWorkEventArgs ea2) =>
            {
                Action exec = updateModel;
                exec.Invoke();
            };
            cc.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs ea) =>
            {
                tc.StopClock();
                if (pf != null && pf.Visible) pf.Close();
                showAuswertungen();
            };
            return cc;
        }


        private void showAuswertungen()
        {
            JosefPane MyJosefPane = Globals.ThisAddIn.MyJosefPane;
            MyJosefPane.checkBox1.CheckedChanged
                += new System.EventHandler(this.checkBox1_CheckedChanged);
            MyJosefPane.ignorierenToolStripMenuItem.Click 
                += new System.EventHandler(this.ignorierenToolStripMenuItem_Click);
            Globals.ThisAddIn.TaskPane.Width = 344;
            Globals.ThisAddIn.TaskPane.Visible = true;
            stats = new JosefStatistics(model);
            setLabels(stats, MyJosefPane.label7, 
                             MyJosefPane.label8, 
                             MyJosefPane.label9, 
                             MyJosefPane.label10, 
                             MyJosefPane.label13);
            
            fillHistogram(MyJosefPane.chart1);
            fillGridViewNgram(MyJosefPane.dataGridView1);
            fillGridViewFreq(MyJosefPane.dataGridView2,
                MyJosefPane.checkBox1.Checked);
            MyJosefPane.setNodes(buildTreeNodes(stats));
            MyJosefPane.treeView1.NodeMouseClick +=
            (object sender, System.Windows.Forms.TreeNodeMouseClickEventArgs e) =>
            {
                SentenceTreeNode stn = (SentenceTreeNode)e.Node;
                model.Document.Range(stn.begin, stn.end).Select();
                
            };
        }

        
        private void ignorierenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ContextMenuStrip cms = (ContextMenuStrip)((ToolStripMenuItem)sender).Owner;
            TreeView treeView = (TreeView)cms.SourceControl;
            SentenceTreeNode stn =
                (SentenceTreeNode)treeView.GetNodeAt(treeView.PointToClient(cms.Location));
            if (stn.Parent != null)
            {
                switch (stn.Parent.Text.Substring(0,2))
                {
                    case "Fü":
                        stats.fuellwortSaetze
                            .Remove(stats.fuellwortSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                    case "La":
                        stats.langeSaetze
                            .Remove(stats.langeSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                    case "Ko":
                        stats.komplexeSaetze
                            .Remove(stats.komplexeSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                    case "Pa":
                        stats.passivSaetze
                            .Remove(stats.passivSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                    case "Un":
                        stats.unpersoenlicheSaetze
                            .Remove(stats.unpersoenlicheSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                    case "Ph":
                        stats.phrasenSaetze
                            .Remove(stats.phrasenSaetze
                            .Where(s => s.Start == stn.begin && s.End == stn.end).First());
                        break;
                }
                Globals.ThisAddIn.MyJosefPane.setNodes(buildTreeNodes(stats));
            }
        }
                
        void fillHistogram(Chart chart1)
        {

            int[] lengths = model.Sentences.Select(s => s.Words.Count).ToArray();
            for (int i = lengths.Min(); i < lengths.Max() + 1; i++)
            {
                chart1.Series[0].Points.AddXY(i, lengths.Where(l => l == i).Count());
            }
        }

        void setLabels(JosefStatistics stats, Label lbl7, Label lbl8, Label lbl9, Label lbl10, Label lbl13)
        {
            lbl7.Text = "" + stats.BasicStatistics.paragraphCount;
            lbl9.Text = "" + stats.BasicStatistics.wordCount;
            lbl8.Text = "" + stats.BasicStatistics.sentenceCount;
            lbl10.Text = "" + stats.BasicStatistics.uniqueWordCount;
            lbl13.Text = "" + stats.BasicStatistics.dickesSteiwer+"%";
        }

        void fillGridViewFreq(DataGridView dataGridView2, bool filter)
        {
            float wcount = model.Words.Count;
            dataGridView2.Rows.Clear();
            foreach (KeyValuePair<string, int> v in model.UniqueWords.OrderByDescending(s => s.Value).Take(100))
            {
                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
                row.CreateCells(dataGridView2);
                row.Cells[0].Value = v.Key;
                row.Cells[1].Value = v.Value;
                row.Cells[2].Value = Math.Round(100.0f * v.Value / wcount, 2) + "%";
                if (filter)
                {
                    if (!StopWords.stopWords.Contains(v.Key.ToLower()))
                        dataGridView2.Rows.Add(row);
                }
                else
                {
                    dataGridView2.Rows.Add(row);
                }
            }
        }

        void fillGridViewNgram(DataGridView dataGridView1)
        {
            
            dataGridView1.Rows.Clear();
            foreach (KeyValuePair<string, int> v in model.ngrams(2).OrderByDescending(ng => ng.Value))
            {
                System.Windows.Forms.DataGridViewRow row = new System.Windows.Forms.DataGridViewRow();
                row.CreateCells(dataGridView1);
                row.Cells[0].Value = v.Key;
                row.Cells[1].Value = v.Value;
                dataGridView1.Rows.Add(row);
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
            List<SentenceTreeNode> unpersNodes = new List<SentenceTreeNode>();
            List<SentenceTreeNode> phrasenNodes = new List<SentenceTreeNode>();
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
            foreach (Word.Range satz in statistics.unpersoenlicheSaetze)
            {
                unpersNodes = addNodeToList(satz, unpersNodes, true);
            }
            foreach (Word.Range satz in statistics.phrasenSaetze)
            {
                phrasenNodes = addNodeToList(satz, phrasenNodes, true);
            }
            return buildParentNodes(statistics.BasicStatistics.sentenceCount,fuellNodes,
                                    langNodes, komplexNodes, passivNodes, unpersNodes, phrasenNodes);
        }

        private SentenceTreeNode[] buildParentNodes(int sc,List<SentenceTreeNode> fuellNodes,
                                                           List<SentenceTreeNode> langNodes,
                                                           List<SentenceTreeNode> komplexNodes,
                                                           List<SentenceTreeNode> passivNodes,
                                                           List<SentenceTreeNode> unpersNodes,
                                                           List<SentenceTreeNode> phrasenNodes)
        {
            string end=" von " + sc + ")";
            SentenceTreeNode fuellnode = new SentenceTreeNode("Füllwörter-Sätze (" + fuellNodes.Count + end, fuellNodes.ToArray()),
                             langnode = new SentenceTreeNode("Lange Sätze (" + langNodes.Count + end, langNodes.ToArray()),
                             komplexnode = new SentenceTreeNode("Komplexe Sätze (" + komplexNodes.Count + end, komplexNodes.ToArray()),
                             passivnode = new SentenceTreeNode("Passiv-Sätze (" + passivNodes.Count + end, passivNodes.ToArray()),
                             unpersnode = new SentenceTreeNode("Unpersönliche Sätze (" + unpersNodes.Count + end, unpersNodes.ToArray()),
                             phrasennode = new SentenceTreeNode("Phrasen-Sätze (" + phrasenNodes.Count + end, phrasenNodes.ToArray()); ;
            fuellnode.BackColor = ((float)fuellNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)fuellNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            langnode.BackColor = ((float)langNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)langNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            komplexnode.BackColor = ((float)komplexNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)komplexNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            passivnode.BackColor = ((float)passivNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)passivNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            unpersnode.BackColor = ((float)unpersNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)unpersNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            phrasennode.BackColor = ((float)phrasenNodes.Count / sc < 0.3f) ? System.Drawing.Color.PaleGreen : (((float)phrasenNodes.Count / sc < 0.7f) ? System.Drawing.Color.NavajoWhite : System.Drawing.Color.Tomato);
            return new SentenceTreeNode[] { fuellnode,
                                            langnode,
                                            komplexnode,
                                            passivnode,
                                            unpersnode,
                                            phrasennode};
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

        void group1_DialogLauncherClick(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            dlg = new OptionsDialog();
            DialogResult result = dlg.ShowDialog();
            dlg.FormClosed += dlg_FormClosed;
        }


        void dlg_FormClosed(object sender, FormClosedEventArgs e)
        {
            //TODO: Do stuff here!!
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            DerPapierjosef.View.AboutBox about = new DerPapierjosef.View.AboutBox();
            about.ShowDialog();
        }
    }
}
