#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AllowLatinTextWrapInMiddleOfAWord.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AllowLatinTextWrapInMiddleOfAWord.docx";

	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Paragraph> para = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);
	
	para->GetFormat()->SetWordWrap(false);
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
