#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";
	wstring outputFile = OUTPUTPATH"AddDateTimeStamp.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(0);

	//Set the font and brush
	LPCWSTR_S arial = L"Arial";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(arial, 12.f, PdfFontStyle::Regular, true);

	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetBlue()));

	//Time text
	LPCWSTR_S t = L"MM/dd/yy hh:mm:ss tt";
	wstring timeString = DateTime::GetNow()->ToString(t);

	//Create a template and rectangle, draw the string
	intrusive_ptr<PdfTemplate> template1 = new PdfTemplate(140, 15);
	intrusive_ptr<RectangleF> rect = new RectangleF(new PointF(page->GetActualSize()->GetWidth() - template1->GetWidth() - 10,
		page->GetActualSize()->GetHeight() - template1->GetHeight() - 10), template1->GetSize());
	template1->GetGraphics()->DrawString(timeString.c_str(), font, brush, new PointF(0, 0));

	//Create stamp annoation
	intrusive_ptr<PdfRubberStampAnnotation> stamp = new PdfRubberStampAnnotation(rect);
	intrusive_ptr<PdfAppearance> apprearance = new PdfAppearance(stamp);
	apprearance->SetNormal(template1);
	stamp->SetAppearance(apprearance);
	page->GetAnnotationsWidget()->Add(stamp);

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
