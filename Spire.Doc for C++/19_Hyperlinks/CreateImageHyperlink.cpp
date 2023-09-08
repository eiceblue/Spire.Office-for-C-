#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BlankTemplate.docx";
	wstring inputFile_1 = input_path + L"Spire.Doc.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateImageHyperlink.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = doc->GetSections()->GetItemInSectionCollection(0);
	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	//Load an image to a DocPicture object
	intrusive_ptr<DocPicture> picture = new DocPicture(doc);
	//Add an image hyperlink to the paragraph
	picture->LoadImageSpire(inputFile_1.c_str());

	paragraph->AppendHyperlink(L"https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType::WebLink);
	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
