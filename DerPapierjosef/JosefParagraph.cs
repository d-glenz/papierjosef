using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DerPapierjosef
{
    
    class JosefParagraph
    {
        bool ignore;
        JosefSentence[] sentences;
        List<String> words;
        Word.Paragraph paragraph;
        Word.Document document;

        public bool Ignore
        {
            get
            {
                return ignore;
            }
        }

        public JosefSentence[] Sentences
        {
            get
            {
                return sentences;
            }
        }

        public List<String> Words
        {
            get
            {
                return words;
            }
        }

        public Word.Paragraph Paragraph
        {
            get
            {
                return paragraph;
            }
        }

        public JosefParagraph(Word.Paragraph p,Word.Document d,OpenNLP nlp)
        {
            paragraph = p;
            document = d;
            words = new List<string>();
            if (paragraph.Range.Text.Count() > 1 && paragraph.get_Style().NameLocal.Contains("Standard"))
            {
                ignore = false;

                string paragraphText = paragraph.Range.Text;
                

                opennlp.tools.util.Span[] spans = nlp.sentence.sentPosDetect(paragraphText);
                int pStart = paragraph.Range.Start;
                sentences = new JosefSentence[spans.Count()];
                for (int i = 0; i < spans.Count(); i++)
                {
                    //paragraph.Range.Sentences[i + 1]
                    sentences[i] = new JosefSentence(document.Range(spans[i].getStart() + pStart, spans[i].getEnd() + pStart), 
                                                     document, 
                                                     nlp);
                    words.AddRange(sentences[i].Words);
                }
            }else{
                ignore=true;
            }
            
        }
    }
}
