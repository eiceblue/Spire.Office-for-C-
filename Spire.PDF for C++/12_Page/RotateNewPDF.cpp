#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"RotateNewPDF.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	//Create PdfUnitConvertor to convert the unit
	intrusive_ptr<PdfUnitConvertor> unitCvtr = new PdfUnitConvertor();

	//Setting for page margin
	intrusive_ptr<PdfMargins> margin = new PdfMargins();
	margin->SetTop(unitCvtr->ConvertUnits(2.54f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetBottom(margin->GetTop());
	margin->SetLeft(unitCvtr->ConvertUnits(2.0f, PdfGraphicsUnit::Centimeter, PdfGraphicsUnit::Point));
	margin->SetRight(margin->GetLeft());

	//Create PdfSection
	intrusive_ptr<PdfSection> section = doc->GetSections()->Add();

	//Set "A4" for Pdf page
	section->GetPageSettings()->SetSize(new SizeF(PdfPageSize::A4()));

	//Set page margin
	section->GetPageSettings()->SetMargins(margin);

	//Set rotating angle
	section->GetPageSettings()->SetRotate(PdfPageRotateAngle::RotateAngle90);

	//Add the page
	intrusive_ptr<PdfPageBase> page = section->GetPages()->Add();

	//Define a PdfBrush
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetBlack();

	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 13.0f, PdfFontStyle::Bold, true);

	//Set the string format 
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Left);

	//Set the position for drawing 
	float x = 0;
	float y = 50;

	//Text string 
	std::wstring specification = L"The sample demonstrates how to rotate page when creating a PDF document.";

	//Draw text string on page canvas
	page->GetCanvas()->DrawString(specification.c_str(), font, brush, x, y, format);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
