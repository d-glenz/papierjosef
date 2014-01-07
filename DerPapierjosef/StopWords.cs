using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DerPapierjosef
{
    class StopWords
    {
        public static string[] stopWords = new string[]{"aber", "als", "am", "an", "auch", "auf", "aus", "bei", "bin", "bis", "bist", "da", "dadurch", "daher", "darum", "das", "daß", "dass", "dein",
            "deine", "dem", "den", "der", "des", "dessen", "deshalb", "die", "dies", "dieser", "dieses", "doch", "dort", "du", "durch", "ein", "eine", "einem",
            "einen", "einer", "eines", "er", "es", "euer", "eure", "für", "hatte", "hatten", "hattest", "hattet", "hier", "hinter", "ich", "ihr", "ihre",
            "ihm", "ihn", "im", "in", "ist", "ja", "jede", "jedem", "jeden", "jeder", "jedes", "jener", "jenes", "jetzt", "kann", "kannst", "können", "könnt",
            "machen", "mein","meine", "mit", "muß", "mußt", "musst", "müssen", "müßt", "nach", "nachdem", "nein", "nicht", "nun", "oder", "seid", "sein", "seine",
            "sich", "sie","sind", "soll", "sollen", "sollst", "sollt", "sonst", "soweit", "sowie", "und", "unser", "unsere", "unter", "vom", "von", "vor",
            "wann", "war","warum", "was", "weiter", "weitere", "wenn", "wer", "werde", "werden", "werdet", "weshalb", "wie", "wieder", "wieso", "wir", "wird","wirst", 
            "wo","woher", "wohin", "zu", "zum", "zur", "über"};

        public static string[] multiFloskel = new string[] { "halt eben", "an der Stelle", "an dieser Stelle", "im Endeffekt", "ein Stück weit", 
            "gesetzt den Fall", "nicht wirklich", "sage ich jetzt mal", "ich würde meinen", "in der Hinsicht", 
            "in der Situation","vor Ort", "ist ja nicht an dem", "mehr oder minder", "mehr oder weniger", "vom Prinzip", 
            "vom Grundsatz", "vom Ding", "von daher gesehen", "von der Sache her", "ich denke", "wie gesagt", "wenn man so will", 
            "wenn Sie so wollen", "allem Anschein nach", "an sich", "aufs Neue", "bei weitem", "dann und wann", "des Öfteren", 
            "ein bisschen", "ein wenig", "ganz gerne", "ganz und gar", "gar nicht", "hie und da", "im allgemeinen", 
            "im Grunde genommen", "im Prinzip", "in der Tat", "in diesem Zusammenhang", "mehr oder weniger", "meines Erachtens", 
            "ohne weiteres", "ohne Zweifel", "samt und sonders", "sowohl als auch", "üblicher Weise", "voll und ganz", "von neuem" };

        public static string[] singleFloskel = new string[] { "dahingehend", "dementsprechend", "sozusagen", "diesbezüglich", "letztendlich", 
            "letztlich", "endlich", "einfach", "eh", "vollumfänglich", "quasi", "aber", "abermals", "allein", "allemal", "allenfalls",
            "allenthalben", "allerdings", "allesamt", "allzu", "also", "alt", "andauernd", "andererseits", "andernfalls", 
            "anscheinend", "auch", "auffallend", "augenscheinlich", "ausdrücklich", "ausgerechnet", "ausnahmslos", "außerdem", 
            "äußerst", "beinahe", "bekanntlich", "bereits", "besonders", "bestenfalls", "bestimmt", "bloß", "dabei", "dadurch", 
            "dafür", "dagegen", "daher", "damals", "danach", "dann", "dazu", "demgegenüber", "demgemäß", "demnach", "denkbar", 
            "denn", "dennoch", "deshalb", "desungeachtet", "deswegen", "diesmal", "doch", "durchaus", "durchweg", "eben", 
            "eigentlich", "einerseits", "einfach", "einige", "einigermaßen", "einmal", "entsprechend", "ergo", "erheblich", 
            "etliche", "etwa", "etwas", "fast", "folgendermaßen", "folglich", "förmlich", "fortwährend", "fraglos", "freilich", 
            "ganz", "gänzlich", "gar", "gelegentlich", "gemeinhin", "genau", "geradezu", "gewiss", "gewisse", "gewissermaßen", 
            "gewiss", "glatt", "gleichsam", "gleichwohl", "glücklicherweise", "gottseidank", "größtenteils", "halt", "hätte", 
            "häufig", "hingegen", "hinlänglich", "höchst", "höchstens", "immer", "immerhin", "immerzu", "indessen", "infolgedessen", 
            "insbesondere", "insofern", "inzwischen", "irgend", "irgendein", "irgendjemand", "irgendwann", "irgendwas", "irgendwie", 
            "irgendwo", "ja", "je", "jedenfalls", "jedoch", "jemals", "kaum", "keinesfalls", "keineswegs", "könnte", "längst", 
            "lediglich", "leider", "letztendlich", "letztlich", "mal", "manchmal", "mehrfach", "meinetwegen", "meist", "meistens", 
            "meistenteils", "mindestens", "mithin", "mitunter", "möchte", "möglich", "möglicherweise", "möglichst", "nämlich", 
            "naturgemäß", "natürlich", "neuerdings", "neuerlich", "neulich", "nichtsdestoweniger", "nie", "niemals", "nun", "nunmehr",
            "nur", "offenbar", "offenkundig", "offensichtlich", "oft", "ohnedies", "partout", "persönlich", "plötzlich", "praktisch",
            "quasi", "recht", "reichlich", "reiflich", "relativ", "restlos", "richtiggehend", "riesig", "rundheraus", "rundum", 
            "sattsam", "schlicht", "schlichtweg", "schließlich", "schlussendlich", "schlußendlich", "schon", "schwerlich", "sehr", 
            "selbst", "selbstredend", "selbstverständlich", "selten", "seltsamerweise", "sicher", "sicherlich", "so", "sogar", 
            "sonst", "sowieso", "sozusagen", "stellenweise", "stets", "tatsächlich", "trotzdem", "überaus", "überdies", "überhaupt", 
            "übrigens", "umständehalber", "unbedingt", "unbeschreiblich", "unerhört", "ungefähr", "ungemein", "ungewöhnlich", 
            "ungleich", "unglücklicherweise", "unlängst", "unmaßgeblich", "unsagbar", "unsäglich", "unstreitig", "unzweifelhaft", 
            "vergleichsweise", "vermutlich", "vielfach", "vielleicht", "voll", "vollends", "völlig", "vollkommen", "vollständig", 
            "wahrscheinlich", "weidlich", "weitgehend", "wenigstens", "wieder", "wiederum", "wirklich", "wohl", "wohlgemerkt", 
            "womöglich", "ziemlich", "zudem", "zugegeben", "zumeist", "zusehends", "zusehens", "zuweilen", "zweifellos", 
            "zweifelsfrei", "zweifelsohne"};

        public static int fuellwortAnzahl(List<string> words)
        {
            List<string> result = new List<string>();
            foreach (string word in words)
            {
                if (Array.IndexOf(singleFloskel, word) != -1)
                {
                    result.Add(word);
                }
            }
            return result.Count;
        }

        public static int phrasenAnzahl(List<string> words)
        {
            List<string> result = new List<string>();
            foreach (string floskel in StopWords.multiFloskel)
            {
                if (String.Join(" ", words).Contains(floskel))
                    result.Add(floskel);
            }
            return result.Count;
        }
    }
}
