#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DrawingTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DrawDashedLine.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfGraphicsState> state = page->GetCanvas()->Save();
	float x = 150;
	float y = 200;
	float width = 300;
	intrusive_ptr<PdfPen> pen = new PdfPen(new PdfRGBColor(Spire::Common::Color::GetRed()), 3.f);
	pen->SetDashStyle(PdfDashStyle::Dash);
	float flo[] = { 1,4,1 };
	vector<float> vec;
	vec.insert(vec.begin(), flo, flo + 3);
	pen->SetDashPattern(vec);
	page->GetCanvas()->DrawLine(pen, x, y, x + width, y);
	page->GetCanvas()->Restore(state);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

