using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using opennlp;

namespace DerPapierjosef
{
    class JosefSentence
    {
        List<string> words;
        //Word.Range sentence;
        Word.Document document;
        int startPos, endPos;
        List<int> subOrdinatePos;
        public string[] tags;
        public string[] chunks;
        public bool ignore;

        public enum Verstaendlichkeit
        {
            SehrLeicht,
            Leicht,
            Verstaendlich,
            Schwer,
            SehrSchwer
        }

        public Verstaendlichkeit Satzverstaendlichkeit //Reiners (1969)
        {
            get
            {
                return (words.Count<14)?
                    Verstaendlichkeit.SehrLeicht:
                    ((words.Count<19)?Verstaendlichkeit.Leicht:
                        ((words.Count<25)?Verstaendlichkeit.Verstaendlich:
                            ((words.Count<30)?Verstaendlichkeit.Schwer:Verstaendlichkeit.SehrSchwer)));
            }
        }

        public List<string> Words
        {
            get
            {
                return words;
            }
        }

        public Word.Range Sentence
        {
            get
            {
                return document.Range(startPos,endPos);
            }
        }

        public string[] Subordinates
        {
            get
            {
                string[] result=new string[subOrdinatePos.Count];
                for (int i = 0; i < subOrdinatePos.Count-1; i++)
                {
                    result[i] = document.Range(startPos, endPos).Text.Substring(subOrdinatePos[i], subOrdinatePos[i + 1]);
                }
                return result;
            }
        }

        public List<string> FloskelnInSentence
        {
            get
            {
                List<string> result = new List<string>();
                foreach (string word in words)
                {
                    if (Array.IndexOf(StopWords.singleFloskel, word) != -1)
                    {
                        result.Add(word);
                    }
                }
                foreach (string floskel in StopWords.multiFloskel)
                {
                    if (document.Range(startPos, endPos).Text.Contains(floskel))
                        result.Add(floskel);
                }
                return result;
            }
        }

        public float FloskelFaktor
        {
            get
            {
                return (float) FloskelnInSentence.Count / Words.Count;
            }
        }

        public float NominalstilFaktor
        {
            get
            {
                int noNomin = 0;
                bool weakVerb = false;
                for (int i = 0; i < words.Count;i++)
                {
                    if(tags[i] == "NN" &&
                        (words[i].EndsWith("ung") || 
                         words[i].EndsWith("heit") || 
                         words[i].EndsWith("keit") ||
                         words[i].EndsWith("en")))
                    {
                        noNomin++;
                    }
                   
                }

                weakVerb = words.Contains("wird")  //schwaches Verb
                        || words.Contains("ist") 
                        || words.Contains(" hat");

                return ((float)noNomin / words.Count)+(weakVerb?0.2f:0.0f);
            }
        }

        public bool Zahlwortfehler
        {
            get
            {
                Word.Range sentence = document.Range(startPos, endPos);
                return sentence.Text.Contains(" 0 ") || sentence.Text.Contains(" 1 ") || sentence.Text.Contains(" 2 ") ||
                    sentence.Text.Contains(" 3 ") || sentence.Text.Contains(" 4 ") || sentence.Text.Contains(" 5 ") ||
                    sentence.Text.Contains(" 6 ") || sentence.Text.Contains(" 7 ") || sentence.Text.Contains(" 8 ") ||
                    sentence.Text.Contains(" 9 ") || sentence.Text.Contains(" 10 ") || sentence.Text.Contains(" 11 ") ||
                    sentence.Text.Contains(" 12 ");
            }
        }

        public JosefSentence(Word.Range s, Word.Document d, OpenNLP nlp)
        {
            ignore = false;
            startPos = s.Start;
            endPos   = s.End;
            document = d;
            words = new List<string>();
            for (int i = 0; i < s.Words.Count; i++)
            {
                words.Add(s.Words[i + 1].Text);
            }

            subOrdinatePos=new List<int>();
            subOrdinatePos.Add(0);
            //System.Windows.Forms.MessageBox.Show(s.Text);
            while (s.Text.IndexOf(", ",subOrdinatePos.Last()) > 0)
            {
                subOrdinatePos.Add(s.Text.IndexOf(", ", subOrdinatePos.Last()) + 1);
            }
            subOrdinatePos.Add(s.Text.Length);
            tags=nlp.tagger.tag(words.ToArray());
            chunks=nlp.chunker.chunk(words.ToArray(), tags);
        }


    }
}
