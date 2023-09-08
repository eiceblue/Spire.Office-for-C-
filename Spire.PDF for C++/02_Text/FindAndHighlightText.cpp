#include "../pch.h"
#include <algorithm> 

using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"FindAndHighlightText.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile=output_path+L"FindAndHighlightText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	for (int i = 0; i < doc->GetPages()->GetCount(); i++) {
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		for(intrusive_ptr<PdfTextFind> find : page->FindText(L"science", Find_TextFindParameter::None)->GetFinds())
		{
			find->ApplyHighLight();
		}
	}
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

