#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"AddPdfLaunchAction.pdf";
	//Create a new PDF document and add a page to it
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();

	//Create a PDF Launch Action       
	//intrusive_ptr<PdfLaunchAction> launchAction = new PdfLaunchAction(DataPath"Demo/text.txt)");

	wstring inputFile = DATAPATH"text.txt";
	intrusive_ptr<PdfLaunchAction> launchAction = new PdfLaunchAction(outputFile.c_str());
	//Create a PDF Action Annotation with the PDF Launch Action
	wstring text = L"Click here to open file";

	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 13.0f, PdfFontStyle::Regular, true);
	intrusive_ptr<RectangleF> rect = new RectangleF(50, 50, 230, 15);
	page->GetCanvas()->DrawString(text.c_str(), font, PdfBrushes::GetForestGreen(), rect);
	intrusive_ptr<PdfActionAnnotation> annotation = new PdfActionAnnotation(rect, launchAction);

	//Add the PDF Action Annotation to page
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	newPage->GetAnnotations()->Add(annotation);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
