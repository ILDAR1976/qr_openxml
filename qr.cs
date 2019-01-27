using QRCoder;
using Pdf417;
using DataMatrix.net;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Runtime.InteropServices;
using System.Runtime;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Collections.Generic;


class GenerateQR {
	
	private Dictionary<String, String> marks { get; set; }
	private static string srcfile = @".\temp.docx";
	private static string file = @".\job.docx";
	private static string imageFile = @".\pdf417.bmp";
	private static string labelText = "[Dash_code]";
	private static DataMatrix.net.DmtxImageEncoder DataEncoder;
    private static DataMatrix.net.DmtxImageEncoderOptions DataEncodeOptions;
    private static Image Dataimg;
    private static Bitmap Databitmap;
    private static System.Drawing.Color DM_forecolour = System.Drawing.Color.Black;
    private static System.Drawing.Color DM_backcolour = System.Drawing.Color.White; 
	private static Bitmap barcodeImage;
	 
	public GenerateQR() {
		marks = new Dictionary<String, String>()  {
			{"[First]","первая"},
			{"[Second]","вторая"},
			{"[Third]","третья"},
			{"[Fourth]","четвертая"},
			{"[Fifth]","пятая"}
		};
	}
	
	static void Main(string[] args)	{
		
		GenerateQR gqr = new GenerateQR() ;
		
		generatePDF417("ST00012|Москва");	
		generaterQr("ST00012|Москва");		 
				
		File.Copy(srcfile, file, true);
		  
		using (var document = WordprocessingDocument.Open(file, isEditable: true))
		{
			
			
			DocumentFormat.OpenXml.Wordprocessing.Table table = document.MainDocumentPart.Document.Body.Descendants<Table>().ElementAt(0);
			DocumentFormat.OpenXml.OpenXmlElement  table2 = ((DocumentFormat.OpenXml.OpenXmlElement) table).CloneNode(true);
			
			full_table(	document, table, gqr.marks);
			
			Paragraph p1 = new Paragraph(new Run(new Text("")));
			Paragraph p2 = new Paragraph(new Run(new Text("")));
					
			document.MainDocumentPart.Document.Body.InsertAfter( 
				p1,
				(DocumentFormat.OpenXml.OpenXmlElement) table
			);
		
		
			document.MainDocumentPart.Document.Body.InsertAfter( 
				p2,
				p1
			);
		
			
			generaterQr("ST00012|Казань");
			GenerateBarcode_DataMatrix("ST00012|Казань");
			gqr.marks["[First]"] = "последняя";
			
			full_table(	document, (DocumentFormat.OpenXml.Wordprocessing.Table) table2, gqr.marks);
			
			document.MainDocumentPart.Document.Body.InsertAfter( 
				table2,
				p2
			);
		
			
				
			document.Save();
			
			int iTest = document.MainDocumentPart.Document.Body.Elements<Table>().Count();
			Console.WriteLine("" + iTest);
		}
	}

	private static void full_block(
		DocumentFormat.OpenXml.Packaging.WordprocessingDocument document,
		Dictionary<String, String> marks) {
		
		
		int table_count = 0;
		foreach (var table in document.MainDocumentPart.Document.Body.Descendants<Table>()) {
			Console.WriteLine("table " + ++table_count);
			
			full_table(document,table,marks);
		}
		
	}


	private static void full_table(	
		DocumentFormat.OpenXml.Packaging.WordprocessingDocument document,
		DocumentFormat.OpenXml.Wordprocessing.Table table,
		Dictionary<String, String> marks) {
	 
		
		//int cell_count = 0;
		
		foreach (var cell in table.Descendants<TableCell>()) {
			//Console.WriteLine(" " + ++cell_count + " cell value: " + cell);
			
			foreach (var item in cell) {

				if (item.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Table)) {
					full_table(document,(DocumentFormat.OpenXml.Wordprocessing.Table) item, marks);
				} else if (item.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Paragraph)) {
							foreach (var mrk in marks) {
								if (cell.InnerText.Contains(labelText)) {
									ImagePart imagePart = document.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
									using (FileStream stream = new FileStream(imageFile, FileMode.Open))
									{
										imagePart.FeedData(stream);
									}
									cell.RemoveAllChildren();
									addImageToCell(cell, document.MainDocumentPart.GetIdOfPart(imagePart));									
								} else if (cell.InnerText.Contains(mrk.Key)) {
									String str = cell.InnerText;
									((DocumentFormat.OpenXml.Wordprocessing.Paragraph) item).RemoveAllChildren();
									((DocumentFormat.OpenXml.Wordprocessing.Paragraph) item).Append(new Run(new Text(str.Replace(mrk.Key,mrk.Value))));
								}
								
							}	
				}
				
			}
		}
		
	}

	private static void generaterQr(String str) {
		QRCodeGenerator.ECCLevel eccLevel = QRCodeGenerator.ECCLevel.L;
		var generator = new QRCodeGenerator();
		var data = generator.CreateQrCode(str, eccLevel);
		int pixelsPerModule = 20;
		string foreground = "#000000";
		string background = "#FFFFFF";

		using (var code = new QRCode(data)) {
			using (var bitmap = code.GetGraphic(pixelsPerModule, foreground, background, true))
			{
				bitmap.Save("qr.png", ImageFormat.Jpeg);
			}
		}	
	}
	
	
	 private static void GenerateBarcode_DataMatrix(string inputData)
        {

            // 
            try // Encode
                {
                    DataEncoder = new DataMatrix.net.DmtxImageEncoder();
                    DataEncodeOptions = new DataMatrix.net.DmtxImageEncoderOptions();

                DataEncodeOptions.ForeColor = DM_forecolour; // Set Fore Color 
                DataEncodeOptions.BackColor = DM_backcolour; // Set Bg Color

                // Encode Data Matrix Image
                Dataimg = DataEncoder.EncodeImage(inputData, DataEncodeOptions);

                // Init Barcode Image
                Databitmap = new Bitmap(Dataimg);
                barcodeImage = Databitmap;

                //Label12.Text = "Hashcode: " + DataEncoder.GetHashCode;
				Databitmap.Save(@".\dm.jpg");
                
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                    return;
                }
               
            }

       

        
	
	private static void generatePDF417(String str) {


            var barcode = new Barcode("6273917032349234", Pdf417.Settings.Default);

            barcode.Canvas.SaveBmp(@".\pdf417.bmp");
	}
	
	private static void addImageToCell(TableCell cell, string relationshipId) {
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
						  Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
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
					  )
					  { Preset = A.ShapeTypeValues.Rectangle }))
			  )
			  { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
		  )
		  {
			DistanceFromTop = (UInt32Value)0U,
			DistanceFromBottom = (UInt32Value)0U,
			DistanceFromLeft = (UInt32Value)0U,
			DistanceFromRight = (UInt32Value)0U
		  });

	    cell.Append(new Paragraph(new Run(element)));
	}	
	
}