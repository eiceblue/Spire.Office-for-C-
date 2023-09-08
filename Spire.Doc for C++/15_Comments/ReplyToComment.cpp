#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Comment.docx";
	wstring imagePath = input_path + L"logo.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplyToComment.docx";

	//Load the document from disk.
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//get the first comment.
	intrusive_ptr<Comment> comment1 = doc->GetComments()->GetItem(0);

	//create a new comment and specify the author and content.
	intrusive_ptr<Comment> replyComment1 = new Comment(doc);
	replyComment1->GetFormat()->SetAuthor(L"E-iceblue");
	replyComment1->GetBody()->AddParagraph()->AppendText(L"Spire.Doc is a professional Word library on operating Word documents.");

	//add the new comment as a reply to the selected comment.
	comment1->ReplyToComment(replyComment1);

	intrusive_ptr<DocPicture> docPicture = new DocPicture(doc);

	docPicture->LoadImageSpire(imagePath.c_str());

	//insert a picture in the comment
	replyComment1->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0)->GetChildObjects()->Add(docPicture);

	//Save the document.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
