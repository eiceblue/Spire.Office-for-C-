#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"AddImageStamp.pdf";
	wstring inputFile = DATAPATH"AddImageStamp.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(0);
	//Create a rubber stamp annotation
	intrusive_ptr<PdfRubberStampAnnotation> loStamp = new PdfRubberStampAnnotation(new RectangleF(new PointF(0, 0), new SizeF(60, 60)));

	//Create an instance of PdfAppearance 
	intrusive_ptr<PdfAppearance> loApprearance = new PdfAppearance(loStamp);
	wstring inputFile_stamp = DATAPATH"image stamp.jpg";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile_stamp.c_str());
	intrusive_ptr<PdfTemplate> template_Keyword = new PdfTemplate(210, 210);

	//Draw a pdf image into pdf template
	template_Keyword->GetGraphics()->DrawImage(image, 60, 60);
	loApprearance->SetNormal(template_Keyword);
	loStamp->SetAppearance(loApprearance);

	//Add the rubber stamp annotation into pdf
	page->GetAnnotationsWidget()->Add(loStamp);

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}

