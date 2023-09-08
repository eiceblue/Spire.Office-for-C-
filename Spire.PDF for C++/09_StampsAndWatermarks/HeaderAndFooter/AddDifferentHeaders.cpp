#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{

	wstring outputFile = OUTPUTPATH"AddDifferentHeaders.pdf";
	wstring inputFile = DATAPATH"MultipagePDF.pdf";
	
	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	wstring header1 = L"Header 1";
	wstring header2 = L"Header 2";

	//Define style
	LPCWSTR_S arial = L"Arial";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 13.f, PdfFontStyle::Regular, true);
	intrusive_ptr<PdfBrush> brush = PdfBrushes::GetRed();
	intrusive_ptr<RectangleF> rect = new RectangleF(new PointF(0, 20.f), new SizeF(doc->GetPageSettings()->GetSize()->GetWidth(), 50.0f));
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	format->SetAlignment(PdfTextAlignment::Center);
	doc->GetPages()->GetItem(0)->GetCanvas()->DrawString(header1.c_str(), font, brush, rect, format);
	LPCWSTR_S aleo = L"Aleo";
	font = new PdfTrueTypeFont(aleo, 15.f, PdfFontStyle::Regular, true);
	brush = PdfBrushes::GetBlack();
	format->SetAlignment(PdfTextAlignment::Left);
	doc->GetPages()->GetItem(1)->GetCanvas()->DrawString(header2.c_str(), font, brush, rect, format);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}






