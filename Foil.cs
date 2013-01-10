using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.IO;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
//--------------------------------------------------------------

namespace weShow {
    // ------------------------------------------------------------ 

    public class Foil {
    // ------------------------------------------------------------ 
        String ClassName = "Foil";
        private String foilName;                    // Name der Folie
        //private String chapTitle;                   // Name des Kapitels
        private String noteRtf;                     // Foliennotizen
        private String noteHandoutRtf;              // Notizentext der Folie für Handouts
        public String FoilName {
            get { return (foilName != null) ? foilName : WCONST.UNKNOWN; }
            set { foilName = value; } }        
        public int FoilNr { get; set; }             
        //public String ChapTitle {
        //    get { return (chapTitle != null) ? chapTitle : WCONST.UNKNOWN; }
        //    set { chapTitle = value; } }
        public int ChapNr { get; set; }             // Index des Kapitels in der Show
        public double DurationSec { get; set; }     // Dauer der Anzeige der Folie in Sekunden
        public double Delay { get; set; }           // Verzoegerung 
        public double SpeachTime { get; set; }      // Redezeit für die Folie
        public String NoteRtf {
            get { return (noteRtf != null) ? noteRtf : String.Empty; }
            set { noteRtf = value; }
        }
        public String HandoutNoteRtf {
            get { return (noteHandoutRtf != null) ? noteHandoutRtf : String.Empty; }
            set { noteHandoutRtf = value; }
        }
        public Boolean Active { get; set; }         // Folie ist aktiv
        public Boolean IsShowLegend { get; set; }   // Bildtitel anzeigen?
        public Boolean AutoPlay { get; set; }       // automatisch abspielen
        public Boolean IsPrintable { get; set; }    // drucken ja / nein
        public bool IsPrintFoils { get; set; }
        public bool IsPrintNotes { get; set; }
        public bool IsPrintHandouts { get; set; }
        public List<Layer> Layers = new List<Layer>(); // Array mit Ebenen
        public WFONT FoilsFont = new WFONT(30);
        public WFONT FoilsNoteFont = new WFONT(30);
        public WFONT HandoutNoteFont = new WFONT(16);
        public String UndoAction = "unknown";
        public static Canvas ESCFoil {
            get {
                Canvas ca = new Canvas();
                ca.Width = 1024;
                ca.Height = 768;
                ca.Background = Brushes.Black;
                TextBlock tb = new TextBlock();
                tb.Margin = new Thickness(30);
                tb.Foreground = Brushes.White;
                tb.Text = "Taste ESC: Beenden ...";
                ca.Children.Add(tb);
                return ca;
            }
        }

        #region Hilfetexte
        //--------------------------------------------------------------
        public static String CreateToolTip { get { return "Erstellen|Neue Folie auswählen, erstellen und einfügen."; } }
        public static String Create1TextToolTip { get { return "Erstellen|Folie mit einem Text erstellen und einfügen."; } }
        public static String Create2TextToolTip { get { return "Erstellen|Folie mit Titel und Text erstellen und einfügen."; } }
        public static String Create2TitleToolTip { get { return "Erstellen|Titelfolie erstellen und einfügen."; } }
        public static String CreateTextPicToolTip { get { return "Erstellen|Folie mit Text auf der linken Seite und einem Bild."
            + "rechts daneben erstellen und einfügen."; } }
        public static String CreateTextOnPicToolTip { get { return "Erstellen|Folie mit Text auf einem Bild erstellen und einfügen."; } }
        public static String NameToolTip { get { return "Name|Name der Folie."; } }
        public static String DurationToolTip { get { return "Dauer|Anzeigedauer der Folie in Millisekunden."; } }
        public static String DelayToolTip { get { return "Verzögerung|Zeit bevor die Folie angezeigt wird in Sekunden."; } }
        public static String SpeachTimeToolTip { get { return "Redezeit|Geplante Redezeit für die Folie in Sekunden."; } }
        public static String LegendToolTip { get { return "Bildtitel|Bildtitel der Folie anzeigen / nicht anzeigen."; } }
        public static String ActiveToolTip { get { return "Folie anzeigen|Folie anzeigen / nicht anzeigen in der Präsentation."; } }
        public static String AutoPlayToolTip { get { return "Automatisch weiter|Automatisch nach Ablauf der Dauer zur nächsten Folie springen."; } }
        public String WhatIsPrintableToolTip {
            get {
                String s = "Drucken?|Folie:             "; s += (IsPrintFoils) ? "Ja" + WCONST.CRLF : "Nein" + WCONST.CRLF;
                s += "Notizen:        "; s += (IsPrintNotes) ? "Ja" + WCONST.CRLF : "Nein" + WCONST.CRLF;
                s += "Handzettel:   "; s += (IsPrintHandouts) ? "Ja" + WCONST.CRLF : "Nein" + WCONST.CRLF;
                return s;
            }
        }
        public static String WhichIsNewFoilToolTip {
            get {
                String s = "Neue Folie|";
                switch (WSGLOBAL.FoilDefaultLayout) {
                    case "1Text": return s + "Folie mit einem Text erstellen und einfügen.";
                    case "2Text": return s + "Folie mit Titel und Text erstellen und einfügen.";
                    case "2Titel": return s + "Titelfolie erstellen und einfügen.";
                    case "TextLinks": return s + "Folie mit Text auf der linken Seite und einem Bild.";
                    case "TextAufBild": return s + "Folie mit Text auf einem Bild erstellen und einfügen.";
                    default: return s + "Folie mit einem Text erstellen und einfügen.";
                }
            }
        }

        //--------------------------------------------------------------
        #endregion Hilfetexte

        // ------------------------------------------------------------ 

        public Foil(String chapTitle = WCONST.UNKNOWN, String foilName = WCONST.UNKNOWN, String layout = "Kein Layout") {
        // ------------------------------------------------------------
            WERROR.writeLog(ClassName + " " + chapTitle + " " + foilName);
            this.FoilName = foilName;
            this.FoilNr = 1;
            //this.ChapTitle = chapTitle;
            this.DurationSec = 6000;
            this.Delay = 0;
            this.SpeachTime = 2;
            this.Active = true;
            this.IsShowLegend = true;
            this.IsPrintable = true;
            this.IsPrintFoils = true;
            this.IsPrintNotes = true;
            this.IsPrintHandouts = true;
            if (WSGLOBAL.ActShow != null) {
                if (WSGLOBAL.ActShow.FoilsFont != null) FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                if (WSGLOBAL.ActShow.FoilsNoteFont != null) FoilsNoteFont.Get(WSGLOBAL.ActShow.FoilsNoteFont);
                if (WSGLOBAL.ActShow.HandoutNoteFont != null) HandoutNoteFont.Get(WSGLOBAL.ActShow.HandoutNoteFont);
            }

            switch (layout) {
                case "1Text": this.Layers.Add(new Layer()); break;
                case "2Texte": this.Layers.Add(new Layer());
                    Layers[0].IsTitle = true;
                    Layers[0].FoilsFont.Get(WSGLOBAL.ActShow.FoilsTitleFont);
                    Layers[0].BorderHeight.From = 100;
                    Layers[0].ChangeLayerName();
                    this.Layers.Add(new Layer());
                    Layers[1].FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                    Layers[1].BorderHeight.From = 585;
                    Layers[1].Y1.From = 140;
                    Layers[1].LayNr = 1;
                    Layers[1].ChangeLayerName();
                    break;
                case "2Titel": this.Layers.Add(new Layer());
                    Layers[0].IsTitle = true;
                    Layers[0].FoilsFont.Get(WSGLOBAL.ActShow.FoilsTitleFont);
                    Layers[0].FoilsFont.Align = "center";
                    Layers[0].BorderHeight.From = 100;
                    Layers[0].Y1.From = 200;
                    Layers[0].ChangeLayerName();
                    this.Layers.Add(new Layer());
                    Layers[1].FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                    Layers[1].BorderHeight.From = 385;
                    Layers[1].Y1.From = 340;
                    Layers[1].X1.From = 130;
                    Layers[1].BorderWidth.From = 750;
                    Layers[1].LayNr = 1;
                    Layers[1].ChangeLayerName();
                    break;
                case "TextLinks": this.Layers.Add(new Layer());
                    Layers[0].FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                    Layers[0].BorderHeight.From = 690;
                    Layers[0].BorderWidth.From = 300;
                    Layers[0].ChangeLayerName();
                    this.Layers.Add(new Layer(1, "picture"));
                    Layers[1].FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                    Layers[1].X1.From = 340;
                    Layers[1].Y1.From = 30;
                    Layers[1].BorderHeight.From = 690;
                    Layers[1].BorderWidth.From = 660;
                    //Layers[1].LayFile = WSGLOBAL.GetFirstPicture();
                    Layers[1].ChangeLayerName();
                    break;
                case "TextAufBild": this.Layers.Add(new Layer(0, "picture"));
                    Layers[0].BorderHeight.From = 768;
                    Layers[0].BorderWidth.From = 1024;
                    //Layers[0].LayFile = WSGLOBAL.GetFirstPicture();
                    Layers[0].ChangeLayerName();
                    this.Layers.Add(new Layer(1, "text"));
                    Layers[1].FoilsFont.Get(WSGLOBAL.ActShow.FoilsFont);
                    Layers[1].X1.From = 30;
                    Layers[1].Y1.From = 30;
                    Layers[1].BorderHeight.From = 690;
                    Layers[1].BorderWidth.From = 940;
                    Layers[1].ChangeLayerName();
                    break;
            }

            Ellipse ElSound = new Ellipse();
            Ellipse ElSoundBack = new Ellipse();
        }   // ------------------------------------------------------------

        public void Get(Foil startFoil) {
            // --------------------------------------------------------------
            WERROR.writeLog(ClassName + ".getFoil");
            this.FoilName = startFoil.FoilName;
            this.FoilNr = startFoil.FoilNr;
            //this.ChapTitle = startFoil.ChapTitle;
            this.ChapNr = startFoil.ChapNr;
            this.DurationSec = startFoil.DurationSec;
            this.Delay = startFoil.Delay;
            this.SpeachTime = startFoil.SpeachTime;
            this.NoteRtf = startFoil.NoteRtf;
            this.HandoutNoteRtf = startFoil.HandoutNoteRtf;
            this.Active = startFoil.Active;
            this.IsPrintFoils = startFoil.IsPrintFoils;
            this.IsPrintNotes = startFoil.IsPrintNotes;
            this.IsPrintHandouts = startFoil.IsPrintHandouts;
            this.IsShowLegend = startFoil.IsShowLegend;
            FoilsFont.Get(startFoil.FoilsFont);
            FoilsNoteFont.Get(startFoil.FoilsNoteFont);
            HandoutNoteFont.Get(startFoil.HandoutNoteFont);
            this.Layers.Clear();

            for (int i = 0; i < startFoil.Layers.Count; i++) {
                this.Layers.Add(new Layer(i, "text"));
                this.Layers[i].Get(startFoil.Layers[i]);
            }
        }   // --------------------------------------------------------------

        public void UndoPush(String action) {
            // ------------------------------------------------------------  
            Foil tmpFoil = new Foil();
            tmpFoil.Get(this);
            tmpFoil.UndoAction = action;
            WSGLOBAL.UndoFoils.Push(tmpFoil);

            // Kapazität begrenzen
        }   // ------------------------------------------------------------  

        public void RedoPush(String action) {
            // ------------------------------------------------------------  
            Foil tmpFoil = new Foil();
            tmpFoil.Get(this);
            tmpFoil.UndoAction = action;
            WSGLOBAL.RedoFoils.Push(tmpFoil);

            // Kapazität begrenzen
        }   // ------------------------------------------------------------  

        public void ReNumberLayers() {
            // ------------------------------------------------------------  
            for (int i = 0; i < Layers.Count; i++) {
                Layers[i].LayNr = i;
                Layers[i].ChangeLayerName();
            }
        }   // ------------------------------------------------------------  

        public Canvas CreateFoilLayers(bool isLegend = false) {
            // ------------------------------------------------------------
            WERROR.writeLog(ClassName + ".CreateFoilLayers");

            Canvas canFoil = new Canvas();
            canFoil.Background = Brushes.Transparent;
            canFoil.Width = WSGLOBAL.MonitorWidth;
            canFoil.Height = WSGLOBAL.MonitorHeight;

            foreach (Layer la in Layers) {
                    LayerDisplayCanvas tmpCanvas = new LayerDisplayCanvas(la);
                    canFoil.Children.Add(tmpCanvas);
            }

            // Bildunterschrift anzeigen
            if (isLegend) showLegend(canFoil, this.FoilName);

            // Copyright einfügen
            if (FoilNr == 1) showCopyright(canFoil);

            return canFoil;
        }   // -------------------------------------------------------------------

        public void showLegend(Canvas aCan, String fName) {
            // ------------------------------------------------------------ 
            if (FoilNr == 0) return;

            WERROR.writeLog(ClassName + ".showLegend");

            Canvas tmpCanvas = new Canvas();
            Layer tmpLayer = new Layer(Layers.Count, "text");
            tmpLayer.paintString(tmpCanvas, 18, 700, 18, fName);
            aCan.Children.Add(tmpCanvas);
        }   // ------------------------------------------------------------ 

        public void showCopyright(Canvas aCan) {
            // ------------------------------------------------------------ 
            WERROR.writeLog(ClassName + ".showCopyright");

            Canvas crCanvas = new Canvas();
            Layer tmpLayer = new Layer(0, "text");
            tmpLayer.X1.From = 30;
            tmpLayer.Y1.From = 710;
            tmpLayer.FoilsFont.Size = 18;
            tmpLayer.FoilsFont.ForegroundColor = Brushes.Red.Color;
            tmpLayer.FoilsFont.BackgroundColor = Brushes.Transparent.Color;
            tmpLayer.BorderHeight.From = 24;
            String lines = "";
            //int nrLines = 0;
            String cText = "";
            String cMusic = "";
            String cRight = "";

            // Copyrighttext zusammenstellen
            if (WSGLOBAL.ActShow.Chapters.Count > 1 && ChapNr < WSGLOBAL.ActShow.Chapters.Count) {
                cText = WSGLOBAL.ActShow.Chapters[ChapNr].CrText;
                cRight = WSGLOBAL.ActShow.Chapters[ChapNr].Copyright;
                cMusic = WSGLOBAL.ActShow.Chapters[ChapNr].CrMusic;
            }

            if (cText == WCONST.UNKNOWN) cText = "";

            if (cText != "") {
                lines = "Text: " + cText;
                //nrLines = 1;
            }

            if (cMusic == WCONST.UNKNOWN) cMusic = "";

            if (cMusic != "") {
                if (cText != "") lines += " - Musik: " + cMusic;
                else lines = cMusic;
                //nrLines = 1;
            }

            if (cRight == WCONST.UNKNOWN) cRight = "";

            if (cRight != "") {

                if (lines.Length > 0) {
                    lines += WCONST.CRLF + cRight;
                    //nrLines = 2;
                }
                else lines = cRight;
            }

            if (lines.Length > 0) {
                tmpLayer.paintString(crCanvas, 18, 700, 18, lines);
                aCan.Children.Add(crCanvas);
            }
        }   // ------------------------------------------------------------

        public void Save() {
            // ------------------------------------------------------------ 
            // Kapitel speichern
            WERROR.writeLog(ClassName + ".Save");
            WSGLOBAL.ActShow.Chapters[ChapNr].SaveXml(WSGLOBAL.AppActShowFolder);
        }   // ------------------------------------------------------------

        public void Write(XmlWriter xW) {
            // ------------------------------------------------------------ 
            WERROR.writeLog(ClassName + ".Write");

            xW.WriteElementString("foilname", FoilName);
            xW.WriteStartElement("properties");
            xW.WriteAttributeString("duration", DurationSec.ToString());
            xW.WriteAttributeString("delay", Delay.ToString());
            xW.WriteAttributeString("speachtime", SpeachTime.ToString());
            xW.WriteAttributeString("active", Active.ToString());
            xW.WriteAttributeString("isshowlegend", IsShowLegend.ToString());
            xW.WriteAttributeString("autoplay", AutoPlay.ToString());
            xW.WriteEndElement();

            xW.WriteStartElement("print");
            xW.WriteAttributeString("isprintfoils", IsPrintFoils.ToString());
            xW.WriteAttributeString("isprintnotes", IsPrintNotes.ToString());
            xW.WriteAttributeString("isprinthandouts", IsPrintHandouts.ToString());
            xW.WriteEndElement();

            // Schriftangaben zu den Foliennotizen
            FoilsNoteFont.WriteFont(xW, "foilsnotefont");
            FoilsNoteFont.WriteText(xW, "foilsnotetext");

            // Schriftangaben zu den Handzetteltexten
            HandoutNoteFont.WriteFont(xW, "handoutnotefont");
            HandoutNoteFont.WriteText(xW, "handoutnotetext");

            xW.WriteStartElement("notertf");
            xW.WriteCData(NoteRtf);
            xW.WriteEndElement();
            xW.WriteStartElement("handoutnotertf");
            xW.WriteCData(HandoutNoteRtf);
            xW.WriteEndElement();

            int size = this.Layers.Count;

            for (int i = 0; i < size; i++) {
                xW.WriteStartElement("layer");
                Layers[i].Write(xW);
                xW.WriteEndElement();
            }
        }   // ------------------------------------------------------------ 

        public void Read(XmlReader xR) {
            // ------------------------------------------------------------ 
            WERROR.writeLog(ClassName + ".Read Folie: " + FoilNr);
             
            while (xR.Read()) { 
                if (xR.NodeType == XmlNodeType.EndElement) break;
                if (xR.NodeType == XmlNodeType.Element) {
                    switch (xR.Name) {
                        case "foilname": if (xR.MoveToContent() != XmlNodeType.None) FoilName = xR.ReadString(); break;
                        case "duration": if (xR.MoveToContent() != XmlNodeType.None) {
                                DurationSec = Convert.ToInt32(xR.ReadString());
                                if (DurationSec == WCONST.NONUMBER) DurationSec = 0;
                            }
                            break;
                        case "delay": if (xR.MoveToContent() != XmlNodeType.None) {
                                Delay = Convert.ToInt32(xR.ReadString());
                                if (Delay == WCONST.NONUMBER) Delay = 0;
                            }
                            break;
                        case "speachtime": if (xR.MoveToContent() != XmlNodeType.None) {
                                SpeachTime = Convert.ToInt32(xR.ReadString());
                                if (SpeachTime == WCONST.NONUMBER) SpeachTime = 0;
                            }
                            break;
                        case "noticefile": if (xR.MoveToContent() != XmlNodeType.None) { String dummy = xR.ReadString(); } break;
                        case "note": 
                            if (xR.MoveToContent() != XmlNodeType.None) {
                                String note = xR.ReadString();
                                if (note == WCONST.UNKNOWN) note = " ";
                                RichTextBox rtb = new RichTextBox();
                                WRICHTEXT.SetTextAndFonts(note, rtb, FoilsNoteFont);
                                NoteRtf = WRICHTEXT.RtfToString(rtb.Document);
                            }
                            break;
                        case "notertf": xR.Read(); if (xR.MoveToContent() == XmlNodeType.CDATA) { NoteRtf = xR.ReadString(); } break;
                        case "handouttext":
                            if (xR.MoveToContent() != XmlNodeType.None) {
                                String note = xR.ReadString();
                                if (note == WCONST.UNKNOWN) note = " ";
                                RichTextBox rtb = new RichTextBox();
                                WRICHTEXT.SetTextAndFonts(note, rtb, HandoutNoteFont);
                                HandoutNoteRtf = WRICHTEXT.RtfToString(rtb.Document);
                            }
                            break;
                        case "notekeys": if (xR.MoveToContent() != XmlNodeType.None) { String dummy = xR.ReadString(); } break;
                        case "notematerial": if (xR.MoveToContent() != XmlNodeType.None) { String dummy = xR.ReadString(); } break;
                        case "handoutfile": if (xR.MoveToContent() != XmlNodeType.None) { String dummy = xR.ReadString(); } break;
                        case "handoutnotertf": xR.Read(); if (xR.MoveToContent() == XmlNodeType.CDATA) { HandoutNoteRtf = xR.ReadString(); } break;

                        // Folientext
                        case "foilsfont": FoilsFont.ReadFont(xR); break;
                        case "foilstext": FoilsFont.ReadText(xR); break;
                        
                        // Foliennotizen
                        case "foilsnotefont": FoilsNoteFont.ReadFont(xR); break;
                        case "foilsnotetext": FoilsNoteFont.ReadText(xR); break;

                        // Handout
                        case "handoutnotefont": HandoutNoteFont.ReadFont(xR); break;
                        case "handoutnotetext": HandoutNoteFont.ReadText(xR); break;

                        case "active": if (xR.MoveToContent() != XmlNodeType.None) Active = Convert.ToBoolean(xR.ReadString()); break;
                        case "isshowlegend": if (xR.MoveToContent() != XmlNodeType.None) IsShowLegend = Convert.ToBoolean(xR.ReadString()); break;
                        case "autoplay": if (xR.MoveToContent() != XmlNodeType.None) AutoPlay  = Convert.ToBoolean(xR.ReadString()); break;
                        case "isprintable": if (xR.MoveToContent() != XmlNodeType.None) IsPrintable = Convert.ToBoolean(xR.ReadString()); break;
                        case "isprintfoils": if (xR.MoveToContent() != XmlNodeType.None) IsPrintFoils = Convert.ToBoolean(xR.ReadString()); break;
                        case "isprintnotes": if (xR.MoveToContent() != XmlNodeType.None) IsPrintNotes = Convert.ToBoolean(xR.ReadString()); break;
                        case "isprinthandouts": if (xR.MoveToContent() != XmlNodeType.None) IsPrintHandouts = Convert.ToBoolean(xR.ReadString()); break;
                        case "layer":
                            Layer tmpLayer = new Layer();
                            tmpLayer.LayNr = Layers.Count;
                            tmpLayer.Read(xR);
                            tmpLayer.ChangeLayerName();
                            // rtf-Text in Text umwandeln
                            if (WSGLOBAL.ActShow.ShowKind != "book" && tmpLayer.LayText == null) {
                                if (tmpLayer.LayKind == "rtftext") {
                                    FlowDocument fd = WRICHTEXT.StringToRtf(tmpLayer.LayRtfText);
                                    TextRange ttr = new TextRange(fd.ContentStart, fd.ContentEnd);
                                    tmpLayer.LayText = ttr.Text;
                                    tmpLayer.LayKind = "text";
                                }
                            }
                            Layers.Add(tmpLayer);
                            break;
                        case "properties": 
                            Active = false; IsShowLegend = false; AutoPlay = false;
                            if (xR.MoveToContent() != XmlNodeType.None) {
                                if (xR.HasAttributes) {
                                    while (xR.MoveToNextAttribute()) {
                                        switch (xR.Name) {
                                            case "duration": DurationSec = Convert.ToInt32(xR.Value);
                                                if (DurationSec == WCONST.NONUMBER) DurationSec = 0;
                                                break;
                                            case "delay": Delay = Convert.ToInt32(xR.Value);
                                                if (Delay == WCONST.NONUMBER) Delay = 0;
                                                break;
                                            case "speachtime": SpeachTime = Convert.ToInt32(xR.Value);
                                                if (SpeachTime == WCONST.NONUMBER) SpeachTime = 0;
                                                break;
                                            case "active": Active = Convert.ToBoolean(xR.Value); break;
                                            case "isshowlegend": IsShowLegend = Convert.ToBoolean(xR.Value); break;
                                            case "autoplay": AutoPlay = Convert.ToBoolean(xR.Value); break;
                                        }
                                    }
                                }
                            }
                            break;
                        case "print":
                            IsPrintFoils = false; IsPrintNotes = false; IsPrintHandouts = false;
                            if (xR.MoveToContent() != XmlNodeType.None) {
                                if (xR.HasAttributes) {
                                    while (xR.MoveToNextAttribute()) {
                                        switch (xR.Name) {
                                            case "isprintfoils": IsPrintFoils = Convert.ToBoolean(xR.Value); break;
                                            case "isprintnotes": IsPrintNotes = Convert.ToBoolean(xR.Value); break;
                                            case "isprinthandouts": IsPrintHandouts = Convert.ToBoolean(xR.Value); break;
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            if (xR.MoveToContent() != XmlNodeType.None)  { 
                                if (xR.Name != "Foil")
                                    WERROR.writeLog("Falsche XML-Zeile in Foile: " + xR.Name + " : " + xR.ReadString(), 0); 
                            } 
                            break;
                    }
                }
            }
        }   // ------------------------------------------------------------ 

        public static void Create(String layout = "1Text") {
            // ------------------------------------------------------------ 
            // Menü Folie erstellen
            WERROR.writeLog("Foil.Create");

            if (WSGLOBAL.CrtChapter >= WSGLOBAL.ActShow.Chapters.Count) return;

            //if (WSGLOBAL.CrtFoil > WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].Foils.Count) return;

            Foil tmpFoil = new Foil(WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].ChapTitle, "Vorlage", layout);
            tmpFoil.FoilNr = WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].Foils.Count;
            tmpFoil.Active = true;
            tmpFoil.ChapNr = WSGLOBAL.CrtChapter;
            tmpFoil.Delay = WSGLOBAL.ActShow.Delay;
            tmpFoil.DurationSec = WSGLOBAL.ActShow.DurationSec;
            if (tmpFoil.FoilNr < 1) tmpFoil.FoilName = "Kapitelhintergrund";
            else tmpFoil.FoilName = WSGLOBAL.CrtChapter.ToString() + "-" + tmpFoil.FoilNr;

            //switch (WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].ChapKind) {
            //    case "song": tmpFoil.FoilName = "Strophe " + tmpFoil.FoilNr; break;
            //    case "foils": tmpFoil.FoilName = "Folie " + tmpFoil.FoilNr; break;
            //    case "pictures": tmpFoil.FoilName = "Bild " + tmpFoil.FoilNr; break;
            //    case "chapter": tmpFoil.FoilName = "Abschnitt " + tmpFoil.FoilNr; break;
            //    case "media": tmpFoil.FoilName = "Video " + tmpFoil.FoilNr; break;
            //}

            tmpFoil.SpeachTime = 6;
            tmpFoil.IsShowLegend = WSGLOBAL.ActShow.ShowLegend;
            WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].Foils.Add(tmpFoil);
            WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].ReNumberFoils();
            WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].createFoilOrderString();
            WSGLOBAL.ActShow.Chapters[WSGLOBAL.CrtChapter].SaveXml(WSGLOBAL.AppActShowFolder);
            WSGLOBAL.CrtFoil = tmpFoil.FoilNr;
            WSGLOBAL.CrtLayer = 0;
        }   // ------------------------------------------------------------ 

        public void Delete(int ChapNr) {
            // ------------------------------------------------------------
            // Folie aus Kapitel löschen
            WERROR.writeLog(ClassName + ".Delete Kapitel " + ChapNr + " Folie " + FoilNr, 0 );
            WSGLOBAL.ActShow.Chapters[ChapNr].Foils.RemoveAt(this.FoilNr);
            WSGLOBAL.ActShow.Chapters[ChapNr].ReNumberFoils();
            WSGLOBAL.ActShow.Chapters[ChapNr].createFoilOrderString();
            WSGLOBAL.ActShow.Chapters[ChapNr].SaveXml(WSGLOBAL.AppActShowFolder);
            if (WSGLOBAL.ActShow.Chapters[ChapNr].Foils.Count < 2) {
                WSGLOBAL.ActShow.Chapters[ChapNr].Delete();
                return;
            }
            WSGLOBAL.CrtFoil = 1;
        }   // ------------------------------------------------------------ 

        public void InsertFoil(int ChapNr, int foNr) {
            // ------------------------------------------------------------  
            Get(WSGLOBAL.CopyTmpFoil);
            FoilNr = foNr;
            WSGLOBAL.ActShow.Chapters[ChapNr].Foils.Insert(foNr, this);
            WSGLOBAL.ActShow.Chapters[ChapNr].ReNumberFoils();
            WSGLOBAL.ActShow.Chapters[ChapNr].createFoilOrderString();
        }   // ------------------------------------------------------------  

    }   // Foil ------------------------------------------------------------ 

}   // Namespace------------------------------------------------------------ 