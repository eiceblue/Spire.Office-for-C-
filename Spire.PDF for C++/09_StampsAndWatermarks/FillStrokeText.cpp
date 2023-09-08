#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"FillStrokeText.pdf";
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";

	//Load the Pdf from disk
	intrusive_ptr<PdfDocument> document = new PdfDocument();
	document->LoadFromFile(inputFile.c_str());
	//Get the first page
	intrusive_ptr<PdfPageBase> page = document->GetPages()->GetItem(0);

	//Define Pdf pen
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Color::GetGray()));

	//Save graphics state
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();

	//Rotate page canvas
	page->GetCanvas()->RotateTransform(-20);

	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	format->SetCharacterSpacing(5.0f);

	//Draw the string on page
	page->GetCanvas()->DrawString(L"E-ICEBLUE", new PdfFont(PdfFontFamily::Helvetica, 45.f), pen, 0, 500.f, format);

	//Restore graphics
	page->GetCanvas()->Restore(state);

	//Save the document
	document->SaveToFile(outputFile.c_str());
	document->Close();
}
