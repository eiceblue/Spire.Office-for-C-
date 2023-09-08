#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"PDFAToPDF.pdf";
	wstring inputFile = DATAPATH"SamplePDFA.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfNewDocument> newDoc = new PdfNewDocument();
	newDoc->SetCompressionLevel(PdfCompressionLevel::None);
	for (int i = 0; i < doc->GetPages()->GetCount(); i++) {
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		intrusive_ptr<SizeF> size = page->GetSize();
		intrusive_ptr<PdfPageBase> p = newDoc->GetPages()->Add(size, new PdfMargins(0));
		page->CreateTemplate()->Draw(p, 0.f, 0.f);
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
