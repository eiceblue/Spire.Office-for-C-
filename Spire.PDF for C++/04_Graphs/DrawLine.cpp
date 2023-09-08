#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawingTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawLine.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	float x = 95;
	float y = 95;
	float width = 400;
	float height = 500;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetBlack()), 0.1f);
	intrusive_ptr<PdfPen> pen1 = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetRed()), 0.1f);
	page->GetCanvas()->DrawRectangle(pen, x, y, width, height);
	page->GetCanvas()->DrawLine(pen1, x, y, x + width, y + height);
	page->GetCanvas()->DrawLine(pen1, x + width, y, x, y + height);
	page->GetCanvas()->Restore(state);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

