using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using Verstaendlichkeit=DerPapierjosef.JosefSentence.Verstaendlichkeit;

namespace DerPapierjosef
{
    class JosefModel
    {
        JosefParagraph[] paragraphs;
        List<string> words;
        Word.Document document;

        public Word.Document Document
        {
            get
            {
                return document;
            }
        }

        public JosefParagraph[] Paragraphs
        {
            get
            {
                return paragraphs.Where(d => !d.Ignore).ToArray();
            }
        }

        public List<string> Words
        {
            get
            {
                return words;
            }
        }

        public List<string> CleansedWords
        {
            get
            {
                string[] filter = { ",", ".", "(", ")", "-", ").", "" };
                return words.Select(w => w.Trim()).Where(w => !filter.Contains(w)).ToList();
            }
        }

        public List<JosefSentence> Sentences
        {
            get{
                return paragraphs.Where(d => !d.Ignore).SelectMany(d => d.Sentences).ToList();
            }
        }

        public char[] Letters
        {
            get
            {
                return Words.SelectMany(d => d.ToCharArray()).ToArray();
            }
        }

        public Dictionary<string, int> UniqueWords
        {
            get{
                string[] filter = {",",".","(",")","-",").","" };
                return words.Select(w => w.Trim()).Where(w => !filter.Contains(w))
                    .GroupBy(x => x).ToDictionary(x => x.Key, x => x.Count());
            }
        }

        public Dictionary<string, int> FrequentTags
        {
            get
            {
                return Sentences.SelectMany(s => s.tags).GroupBy(x => x).ToDictionary(x => x.Key, x => x.Count());
            }
        }

        

        public JosefModel(Microsoft.Office.Interop.Word.Document doc,OpenNLP nlp,BackgroundWorker bw)
        {
            document = doc;
            foreach (Word.Field f in document.Fields)
            {
                f.Unlink();
            }
            bw.ReportProgress(20);
            paragraphs=new JosefParagraph[document.Paragraphs.Count];

            words = new List<string>();
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                paragraphs[i] = new JosefParagraph(document.Paragraphs[i + 1],doc, nlp);
                words.AddRange(paragraphs[i].Words);
                bw.ReportProgress((int)(80.0f*i / document.Paragraphs.Count)  + 20);
            }
        }

        public JosefSentence[] SentencesByLength(JosefSentence.Verstaendlichkeit v)
        {
            return Sentences.Where(s => s.Satzverstaendlichkeit == v).ToArray();
        }

        
        public Tuple<string,string>[] POSFirstSentence()
        {
            Tuple<string, string>[] result = new Tuple<string, string>[Sentences[1].Words.Count];
            for (int i = 0; i < Sentences[1].Words.Count; i++)
            {
                result[i] = new Tuple<string, string>(Sentences[1].Words[i], Sentences[1].tags[i]);
            }
            return result;
        }

        public Tuple<string, string, string>[] ChunkFirstSentence()
        {

            Tuple<string, string, string>[] result = new Tuple<string, string, string>[Sentences[0].Words.Count];
            for (int i = 0; i < Sentences[0].Words.Count; i++)
            {
                result[i] = new Tuple<string, string, string>(Sentences[0].Words[i], 
                                                              Sentences[0].tags[i], 
                                                              Sentences[0].chunks[i]);
            }
            return result;
        }

        public double Percentile(int[] sequence, float excelPercentile)
        {
            if (sequence.Count() > 0)
            {
                Array.Sort(sequence);
                int N = sequence.Length;
                double n = (N - 1) * excelPercentile + 1;
                // Another method: double n = (N + 1) * excelPercentile;
                if (n == 1d) return sequence[0];
                else if (n == N) return sequence[N - 1];
                else
                {
                    int k = (int)n;
                    double d = n - k;
                    return sequence[k - 1] + d * (sequence[k] - sequence[k - 1]);
                }
            }
            return 0;
        }

        public Dictionary<string, int> ngrams(int n)
        {
            if(n==1)
                return words.GroupBy(x => x).ToDictionary(x => x.Key, x => x.Count());
            else
            {
                Dictionary<string,int> myDict=new Dictionary<string,int>();
                for(int i=0;i<words.Count-1;i++){
                    string bigram=string.Join(" ", words.GetRange(i, 2));
                    if (myDict.ContainsKey(bigram))
                        myDict[bigram] = myDict[bigram] + 1;
                    else
                        myDict[bigram] = 1;
                }
                return myDict;
            }
        }
    }
}
