using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DerPapierjosef
{
    public delegate bool factDetector(List<string> words, List<string> tags);
    class JosefStatistics
    {

        public List<Word.Range> korrelatSaetze;
        public List<Word.Range> passivSaetze;
        public List<Word.Range> vergangenheitsSätze;
        public List<Word.Range> unpersoenlicheSaetze;
        public List<Word.Range> komplexeSaetze;
        public List<Word.Range> langeSaetze;
        public List<Word.Range> phrasenSaetze;
        public List<Word.Range> fuellwortSaetze;
        public List<Word.Range> nominalSaetze;

        private List<Tuple<factDetector, List<Word.Range>>> ruleLogics;

        public JosefStatistics(JosefModel model)
        {
            korrelatSaetze=new List<Word.Range>();
            passivSaetze=new List<Word.Range>();
            vergangenheitsSätze=new List<Word.Range>();
            unpersoenlicheSaetze=new List<Word.Range>();
            komplexeSaetze=new List<Word.Range>();
            langeSaetze=new List<Word.Range>();
            phrasenSaetze=new List<Word.Range>();
            fuellwortSaetze=new List<Word.Range>();
            nominalSaetze=new List<Word.Range>();

            ruleLogics = new List<Tuple<factDetector, List<Word.Range>>>();

            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => w.Count > 20,langeSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => t.Count(tag => tag == "$,") > 2, komplexeSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => StopWords.fuellwortAnzahl(w) > 0, fuellwortSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => StopWords.phrasenAnzahl(w) > 0, phrasenSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => t.Contains("PIS"), unpersoenlicheSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => idxOf(t,"VAFIN")!=-1 && 
                                                                     (w[idxOf(t,"VAFIN")] ==  "wird" || w[idxOf(t,"VAFIN")].StartsWith("wurd") 
                                                                     || w[idxOf(t,"VAFIN")].StartsWith("werd") || w[idxOf(t,"VAFIN")].StartsWith("würd"))
                                                                     && t.Contains("VVPP"), passivSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => (t.Contains("VAFIN") && t.Contains("VVPP")) ||
                                                                               (t.Contains("VAFIN") && t.Contains("VAPP")) ||
                                                                               t.Contains("VMPP") || t.Contains("VVFIN"), 
                                                                               vergangenheitsSätze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => w.Select(ws => ws.EndsWith("ung") || 
                                                                        ws.EndsWith("heit") || 
                                                                        ws.EndsWith("keit")).Count() > 0, 
                                                                        nominalSaetze));
            ruleLogics.Add(new Tuple<factDetector, List<Word.Range>>((w, t) => w.Select(ws => ws.Trim()).Contains("Es") 
                                                                            && w.Select(ws => ws.Trim()).Contains("dass"), 
                                                                            korrelatSaetze));

            foreach (JosefSentence s in model.Sentences)
            {
                foreach (Tuple<factDetector, List<Word.Range>> t in ruleLogics)
                {
                    if (t.Item1(s.Words, s.tags.ToList()))
                        t.Item2.Add(s.Sentence);
                }
            }

            BasicStatistics = new Stats(model.UniqueWords.Count, model.Words.Count, model.Sentences.Count,
                        (int)(model.Percentile(model.Sentences.Select(s => s.Words.Count).ToArray(), 0.9f)),
                        model.Paragraphs.Count(), DickesSteiwer(model.Letters,model.Words,model.Sentences.Count,
                        model.UniqueWords.Count), FloskelMean(model.Sentences), NominalMean(model.Sentences));
        }

        private int idxOf(List<string> tags, string tag)
        {
            return tags.FindIndex(t => t == tag);
        }

        public float DickesSteiwer(char[] Letters, List<string> Words, int sentences, int uniquewords)
        {
            int letters = Letters.Count(),
            words = Words.Count();
            if (words > 0)
            {
                return (float)Math.Round(235.96 - Math.Abs(Math.Log((letters / words) + 1.0)) * 73.021
                    - Math.Abs(Math.Log((words / sentences) + 1)) * 12.564
                    - Math.Abs(Math.Log((uniquewords / words) + 1)) * 50.003, 2);
            }
            return 0;
            
        }

        public float FloskelMean(List<JosefSentence> Sentences)
        {
            if (Sentences.Count > 0)
            {
                return (float)Math.Round(Sentences.Select(s => s.FloskelFaktor).Average() * 100, 2);
            }
            return 0;
        }

        public float NominalMean(List<JosefSentence> Sentences)
        {
            if (Sentences.Count > 0)
            {
                return (float)Math.Round(Sentences.Select(s => s.NominalstilFaktor).Average() * 100, 2);
            }
            return 0;
           
        }

        public class Stats
        {
            public int uniqueWordCount;
            public int wordCount;
            public int sentenceCount;
            public int hardSentenceCount;
            public int paragraphCount;
            public float dickesSteiwer;
            public float floskelMean;
            public float nominalMean;

            public Stats(int uniqueWordCount,
            int wordCount,
            int sentenceCount,
            int hardSentenceCount,
            int paragraphCount,
            float dickesSteiwer,
            float floskelMean,
            float nominalMean)
            {
                this.uniqueWordCount = uniqueWordCount;
                this.wordCount = wordCount;
                this.sentenceCount = sentenceCount;
                this.hardSentenceCount = hardSentenceCount;
                this.paragraphCount = paragraphCount;
                this.dickesSteiwer = dickesSteiwer;
                this.floskelMean = floskelMean;
                this.nominalMean = nominalMean;
            }
        }

        public Stats BasicStatistics;        
    }
}
