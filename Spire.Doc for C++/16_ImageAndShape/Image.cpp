#include "pch.h"
using namespace Spire::Doc;


void InsertImage(intrusive_ptr<Section> section)
{
	//Add paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
	//Add a image and set its width and height
	intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(DATAPATH"/Spire.Doc.png");

	picture->SetWidth(100);
	picture->SetHeight(100);

	paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetLineSpacing(20.0f);
	intrusive_ptr<TextRange> tr = paragraph->AppendText(L"Spire.Doc for C++ is a professional Word C++ library specifically designed for developers to create, read, write, convert, merge, split, and compare Word documents on any C++ platforms with fastand high - quality performance. ");
	tr->GetCharacterFormat()->SetFontName(L"Arial");
	tr->GetCharacterFormat()->SetFontSize(14);

	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetLineSpacing(20.0f);
	tr = paragraph->AppendText(L"As an independent Word C++ API, Spire.Doc for C++ doesn't need Microsoft Word to be installed on neither the development nor target systems. However, it can incorporate Microsoft Word document creation capabilities into any developers' C++ applications.");
	tr->GetCharacterFormat()->SetFontName(L"Arial");
	tr->GetCharacterFormat()->SetFontSize(14);

}

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Image.docx";


	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Add a seciton
	intrusive_ptr<Section> section = document->AddSection();

	//insert image
	InsertImage(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}

