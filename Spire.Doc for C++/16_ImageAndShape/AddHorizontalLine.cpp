#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddHorizontalLine.docx";

	//Create Word document.
	intrusive_ptr<Document> doc = new Document();
	intrusive_ptr<Section> sec = doc->AddSection();
	intrusive_ptr<Paragraph> para = sec->AddParagraph();
	para->AppendHorizonalLine();

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
