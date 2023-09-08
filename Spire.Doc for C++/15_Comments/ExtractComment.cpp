#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"CommentSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractComment.txt";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	wstring* stringBuilder = new wstring();

	//Traverse all comments
	for (int i = 0; i < doc->GetComments()->GetCount(); i++)
	{
		intrusive_ptr<Comment> comment = doc->GetComments()->GetItem(i);
		for (int j = 0; j < comment->GetBody()->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> p = comment->GetBody()->GetParagraphs()->GetItemInParagraphCollection(j);
			stringBuilder->append(p->GetText());
			stringBuilder->append(L"\n");
		}
	}

	//Save to TXT File and launch it
	wofstream write(outputFile);
	write << stringBuilder->c_str();
	write.close();
	doc->Close();
	delete stringBuilder;
}
