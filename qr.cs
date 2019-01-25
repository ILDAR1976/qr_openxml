using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

class GenerateQR {
	static void Main(string[] args)	{
			
			QRCodeGenerator.ECCLevel eccLevel = QRCodeGenerator.ECCLevel.L;
				var generator = new QRCodeGenerator();
			var data = generator.CreateQrCode("ST00020|12|Проверка связи", eccLevel);
					int pixelsPerModule = 20;
					string foreground = "#000000";
					string background = "#FFFFFF";

			using (var code = new QRCode(data)) {
										using (var bitmap = code.GetGraphic(pixelsPerModule, foreground, background, true))
										{
											bitmap.Save("qr.png", ImageFormat.Jpeg);
										}
									

			}	
				
		  string srcfile = @".\temp.docx";
		  string file = @".\job.docx";
		  
		  string imageFile = @".\qr.png";
		  string labelText = "PersonMainPhoto";
		  File.Copy(srcfile, file, true);
		  
		  using (var document = WordprocessingDocument.Open(file, isEditable: true))
		  {
			var mainPart = document.MainDocumentPart;
			var table = mainPart.Document.Body.Descendants<Table>().First();

			var pictureCell = table.Descendants<TableCell>().First(c => c.InnerText == labelText);

			ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

			using (FileStream stream = new FileStream(imageFile, FileMode.Open))
			{
			  imagePart.FeedData(stream);
			}

			pictureCell.RemoveAllChildren();
			AddImageToCell(pictureCell, mainPart.GetIdOfPart(imagePart));

	
			 Paragraph p1 = new Paragraph(new Run(new Text("")));
			 Paragraph p2 = new Paragraph(new Run(new Text("")));
					
			document.MainDocumentPart.Document.Body.InsertAfter( 
				p1,
				table
			);
		
		
			document.MainDocumentPart.Document.Body.InsertAfter( 
				p2,
				p1
			);
		
		
			document.MainDocumentPart.Document.Body.InsertAfter( 
				table.CloneNode(true),
				p2
			);
		
		
			/*
			document.MainDocumentPart.Document.Body.InsertAfter(tbl,table);
			document.MainDocumentPart.Document.Body.InsertAfter(tbl2,tbl);
			*/
			mainPart.Document.Save();
			
			int iTest = mainPart.Document.Body.Elements<Table>().Count();
			Console.WriteLine("" + iTest);
		  }
	}

	private static void AddImageToCell(TableCell cell, string relationshipId)
	{
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