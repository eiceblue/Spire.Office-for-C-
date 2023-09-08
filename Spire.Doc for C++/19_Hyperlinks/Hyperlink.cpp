#include "pch.h"
using namespace Spire::Doc;

void InsertHyperlink(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();
	paragraph->AppendText(L"Spire.Doc for C++ \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Home page");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Contact US");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"mailto:support@e-iceblue.com", L"support@e-iceblue.com", HyperlinkType::EMailLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Forum");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/forum/", L"www.e-iceblue.com/forum/", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Download Link");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", L"www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Insert Link On Image");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	wstring input_path = DATAPATH;
	wstring imagePath = input_path + L"Spire.Doc.png";
	intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(imagePath.c_str());
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType::WebLink);
}

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Hyperlink.docx";

	//Open a blank word document as template
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();

	//Insert hyperlink
	InsertHyperlink(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}


