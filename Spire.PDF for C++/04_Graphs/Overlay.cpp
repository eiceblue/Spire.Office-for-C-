#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Overlay.pdf";

	intrusive_ptr<PdfDocument> doc1 = new PdfDocument();
	intrusive_ptr<PdfDocument> doc2 = new PdfDocument();
	doc1->LoadFromFile((input_path+L"Overlay1.pdf").c_str());
	doc2->LoadFromFile((input_path + L"Overlay2.pdf").c_str());
	intrusive_ptr<PdfTemplate> templates = doc1->GetPages()->GetItem(0)->CreateTemplate();
	for (int i = 0; i < doc2->GetPages()->GetCount(); i++) {
		intrusive_ptr<PdfPageBase> page = doc2->GetPages()->GetItem(i);
		page->GetCanvas()->SetTransparency(0.25f, 0.25f, PdfBlendMode::Overlay);
		page->GetCanvas()->DrawTemplate(templates, PointF::Empty());
	}
	doc2->SaveToFile(outputFile.c_str());
	doc1->Close();
	doc2->Close();
}

