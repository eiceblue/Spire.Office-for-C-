#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";
	wstring outputFile = OUTPUTPATH"ImageWatermarkSecond.pdf";

	//Load Pdf document from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//Load an image
	wstring inputFile_i = DATAPATH"E-logo.png";
	intrusive_ptr<Image> image = Image::FromFile(inputFile_i.c_str());

	//Insert an image into the first PDF page at specific position
	intrusive_ptr<PdfImage> pdfImage = PdfImage::FromFile(inputFile_i.c_str());

	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PointF> position = new PointF(160, 260);
	page->GetCanvas()->Save();
	page->GetCanvas()->SetTransparency(0.5f, 0.5f, PdfBlendMode::Multiply);
	page->GetCanvas()->DrawImage(pdfImage, position);
	page->GetCanvas()->Restore();

	///Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
