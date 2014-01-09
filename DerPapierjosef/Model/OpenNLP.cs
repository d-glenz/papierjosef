using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DerPapierjosef
{
    class OpenNLP
    {
        public opennlp.tools.postag.POSTagger tagger;
        public opennlp.tools.chunker.Chunker chunker;
        public opennlp.tools.sentdetect.SentenceDetector sentence;
        private string posModelPath, 
            tokenModelPath, 
            sentenceModelPath,
            chunkModelPath;
        public OpenNLP(string path)
        {
            posModelPath = path + "de-pos-maxent.bin";
            tokenModelPath= path + "de-token.bin";
            sentenceModelPath = path + "de-sent.bin";
            chunkModelPath = path + "de-chunker.bin";
            tagger = preparePOSTagger();
            chunker = prepareChunker();
            sentence = prepareSentenceDetector();
        }

        enum POS
        {
            ADJA, //attributives Adjektiv
            ADJD, //adverbiales oder prädikatives Adjektiv
            ADV, //Adverb
            APPR, //Präposition; Zirkumposition links
            APPRART, //Präposition mit Artikel
            APPO, //Postposition
            APZR, //Zirkumposition rechts
            ART, //bestimmter oder unbestimmter Artikel
            CARD, //Kardinalzahl
            FM, //Fremdsprachliches Material
            ITJ, //Interjektion
            KOUI, //unterordnende Konjunktion mit ``zu'' und Infinitiv
            KOUS, //unterordnende Konjunktion mit Satz
            KON, //nebenordnende Konjunktion
            KOKOM, //Vergleichskonjunktion
            NN, //normales Nomen
            NE, //Eigennamen
            PDS, //substituierendes Demonstrativpronomen
            PDAT, //attribuierendes Demonstrativpronomen
            PIS, //substituierendes Indefinitpronomen
            PIAT, //attribuierendes Indefinitpronomen ohne Determiner
            PIDAT, //attribuierendes Indefinitpronomen mit Determiner
            PPER, //irreflexives Personalpronomen
            PPOSS, //substituierendes Possessivpronomen
            PPOSAT, //attribuierendes Possessivpronomen
            PRELS, //substituierendes Relativpronomen
            PRELAT, //attribuierendes Relativpronomen
            PRF, //reflexives Personalpronomen
            PWS, //substituierendes Interrogativpronomen
            PWAT, //attribuierendes Interrogativpronomen
            PWAV, //adverbiales Interrogativ- oder Relativpronomen
            PAV, //Pronominaladverb
            PTKZU, //``zu'' vor Infinitiv
            PTKNEG, //Negationspartikel
            PTKVZ, //abgetrennter Verbzusatz
            PTKANT, //Antwortpartikel
            PTKA, //Partikel bei Adjektiv oder Adverb
            TRUNC, //Kompositions-Erstglied
            VVFIN, //finites Verb, voll
            VVIMP, //Imperativ, voll
            VVINF, //Infinitiv, voll
            VVIZU, //Infinitiv mit ``zu'', voll
            VVPP, //Partizip Perfekt, voll
            VAFIN, //finites Verb, aux
            VAIMP, //Imperativ, aux
            VAINF, //Infinitiv, aux
            VAPP, //Partizip Perfekt, aux
            VMFIN, //finites Verb, modal
            VMINF, //Infinitiv, modal
            VMPP, //Partizip Perfekt, modal
            XY, //Nichtwort, Sonderzeichen enthaltend
            KOMMA, //Komma
            SATZENDE, //Satzbeendende Interpunktion
            SATZZEICHEN, //sonstige Satzzeichen; satzintern
        }

#region private methods
        private opennlp.tools.postag.POSTagger preparePOSTagger()
        {
            java.io.FileInputStream tokenInputStream = new java.io.FileInputStream(posModelPath);     //load the token model into a stream
            opennlp.tools.postag.POSModel posModel = new opennlp.tools.postag.POSModel(tokenInputStream); //load the token model
            return new opennlp.tools.postag.POSTaggerME(posModel);  //create the tokenizer
        }
        private opennlp.tools.tokenize.TokenizerME prepareTokenizer()
        {
            java.io.FileInputStream tokenInputStream = new java.io.FileInputStream(tokenModelPath);     //load the token model into a stream
            opennlp.tools.tokenize.TokenizerModel tokenModel = new opennlp.tools.tokenize.TokenizerModel(tokenInputStream); //load the token model
            return new opennlp.tools.tokenize.TokenizerME(tokenModel);  //create the tokenizer
        }
        private opennlp.tools.sentdetect.SentenceDetectorME prepareSentenceDetector()
        {
            java.io.FileInputStream sentModelStream = new java.io.FileInputStream(sentenceModelPath);       //load the sentence model into a stream
            opennlp.tools.sentdetect.SentenceModel sentModel = new opennlp.tools.sentdetect.SentenceModel(sentModelStream);// load the model
            return new opennlp.tools.sentdetect.SentenceDetectorME(sentModel); //create sentence detector
        }
        private opennlp.tools.chunker.ChunkerME prepareChunker()
        {
            java.io.FileInputStream chunkModelStream = new java.io.FileInputStream(chunkModelPath);       //load the sentence model into a stream
            opennlp.tools.chunker.ChunkerModel chunkModel = new opennlp.tools.chunker.ChunkerModel(chunkModelStream);// load the model
            return new opennlp.tools.chunker.ChunkerME(chunkModel); //create sentence detector
        }
        //private opennlp.tools.namefind.NameFinderME prepareNameFinder()
        //{
        //    java.io.FileInputStream modelInputStream = new java.io.FileInputStream(nameFinderModelPath); //load the name model into a stream
        //    opennlp.tools.namefind.TokenNameFinderModel model = new opennlp.tools.namefind.TokenNameFinderModel(modelInputStream); //load the model
        //    return new opennlp.tools.namefind.NameFinderME(model);                   //create the namefinder
        //}
#endregion   
    }
}
