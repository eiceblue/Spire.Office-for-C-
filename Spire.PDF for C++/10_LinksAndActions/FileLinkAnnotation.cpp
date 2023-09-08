#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

void AddFileLinkAnnotation(intrusive_ptr<PdfPageBase> page, float y)
{
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.0f, PdfFontStyle::Regular, true);
	//Set the string format 
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Left);

	//Text string
	std::wstring prompt = L"Launch a File: ";

	//Draw text string on page canvas
	page->GetCanvas()->DrawString(prompt.c_str(), font, PdfBrushes::GetDodgerBlue(), 0, y);

	//Use MeasureString to get the width of string
	float x = font->MeasureString(prompt.c_str(), format)->GetWidth();

	//String of file name
	std::wstring label = L"Sample.pdf";

	//Use MeasureString to get the SizeF of string
	intrusive_ptr<SizeF> size = font->MeasureString(label.c_str());

	//Create a rectangle
	intrusive_ptr<RectangleF> bounds = new RectangleF(x, y, size->GetWidth(), size->GetHeight());

	//Draw label string
	page->GetCanvas()->DrawString(label.c_str(), font, PdfBrushes::GetOrangeRed(), x, y);

	//Create PdfFileLinkAnnotation on the rectangle and link file "Sample.pdf"
	//intrusive_ptr<PdfFileLinkAnnotation> annotation = new PdfFileLinkAnnotation(bounds, DataPath"Demo/Sample.pdf)");

	wstring inputFile = DATAPATH"Sample.pdf";
	intrusive_ptr<PdfFileLinkAnnotation> annotation = new PdfFileLinkAnnotation(bounds, inputFile.c_str());
	//Set color for annotation
	annotation->SetColor(new PdfRGBColor(Color::GetBlue()));

	//Add annotation to the page
	intrusive_ptr<PdfNewPage> newPage = Object::Convert<PdfNewPage>(page);
	newPage->GetAnnotations()->Add(annotation);

}
int main()
{
	wstring outputFile = OUTPUTPATH"FileLinkAnnotation.pdf";
	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Create PdfUnitConvertor to convert the unit
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();

	//Setting for page margin
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(3.0f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	//Add one page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(new SizeF(PdfPageSize::A4()), margin);

	//Define a PdfBrush
	intrusive_ptr<PdfBrush> brush1 = PdfBrushes::GetBlack();

	//Define a font
	intrusive_ptr<PdfTrueTypeFont> font1 = new PdfTrueTypeFont(L"Arial", 13.0f, PdfFontStyle::Bold, true);
	//Set the string format 
	intrusive_ptr<PdfStringFormat> format1 = new PdfStringFormat(PdfTextAlignment::Left);

	//Set the position for drawing 
	float x = 0;
	float y = 50;

	//Text string 
	std::wstring specification = L"The sample demonstrates how to create a file link in PDF document.";

	//Draw text string on page canvas
	page->GetCanvas()->DrawString(specification.c_str(), font1, brush1, x, y, format1);

	//Use MeasureString to get the height of string
	y = y + font1->MeasureString(specification.c_str(), format1)->GetHeight() + 10;

	//Add file link annotation
	AddFileLinkAnnotation(page, y);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
