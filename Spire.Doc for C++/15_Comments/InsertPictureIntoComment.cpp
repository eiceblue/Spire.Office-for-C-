#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"CommentTemplate.docx";
	wstring imagePath = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertPictureIntoComment.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first paragraph and insert comment
	intrusive_ptr<Paragraph> paragraph = doc->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(2);
	intrusive_ptr<Comment> comment = paragraph->AppendComment(L"This is a comment.");
	comment->GetFormat()->SetAuthor(L"E-iceblue");

	//Load a picture
	intrusive_ptr<DocPicture> docPicture = new DocPicture(doc);
	docPicture->LoadImageSpire(imagePath.c_str());

	//Insert the picture into the comment GetBody()
	comment->GetBody()->AddParagraph()->GetChildObjects()->Add(docPicture);

	//Save and launch
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}