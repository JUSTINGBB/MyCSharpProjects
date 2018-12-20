using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace functionTest
{
    class WordHelper_OpenXml
    {
        private const string V = "";

        //创建一个word文档
        public static void CreateWordprocessingDocument(string filepath)
        {
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));

            }
        }
        
        //word创建并添加字符样式
        // Create a new character style with the specified style id, style name and aliases and 
        // add it to the specified style definitions part.
        public static void CreateAndAddCharacterStyle(StyleDefinitionsPart styleDefinitionsPart,
            string styleid, string stylename, string aliases = V)
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new character style and specify some of the attributes.
            Style style = new Style()
            {
                Type = StyleValues.Character,
                StyleId = styleid,
                CustomStyle = true
            };

            // Create and add the child elements (properties of the style).
            Aliases aliases1 = new Aliases() { Val = aliases };
            StyleName styleName1 = new StyleName() { Val = stylename };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "OverdueAmountPara" };
            if (aliases != "")
                style.Append(aliases1);
            style.Append(styleName1);
            style.Append(linkedStyle1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Tahoma" };
            Italic italic1 = new Italic();
            // Specify a 24 point size.
            FontSize fontSize1 = new FontSize() { Val = "48" };
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }

        //word插入图片
        public static void InsertAPicture(string document, string fileName)
        {
            string imgType = fileName.Split('.')[fileName.Split('.').Length - 1];
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                try
                {
                    MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                    ImagePart imagePart = null;
                    //判断图片的格式
                    if (imgType.ToUpper() == "JPEG" || imgType.ToUpper() == "JPE")
                    {
                        imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    }
                    else if (imgType.ToUpper() == "PNG")
                    {
                        imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    }
                    else if (imgType.ToUpper() == "GIF")
                    {
                        imagePart = mainPart.AddImagePart(ImagePartType.Gif);
                    }
                    else if (imgType.ToUpper() == "TIFF" || imgType.ToUpper() == "TIF")
                    {
                        imagePart = mainPart.AddImagePart(ImagePartType.Tiff);
                    }

                    if (imagePart != null)
                    {
                        if (File.Exists(fileName))
                        {
                            using (FileStream stream = new FileStream(fileName, FileMode.Open))
                            {

                                imagePart.FeedData(stream);
                            }

                            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
                        }
                        else
                        {
                            Console.WriteLine("图片文件不存在");
                        }

                    }
                    else
                    {
                        Console.WriteLine("不支持的图片类型");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("文档被占用！"+e);
                }
                
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         ) { Preset = A.ShapeTypeValues.Rectangle }))
                             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}
