#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Comments.pptx";
	wstring outputFile = OUTPUTPATH"GetSlideComments.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Loop through comments
	for (int i = 0; i < ppt->GetCommentAuthors()->GetCount(); i++)
	{
		intrusive_ptr<ICommentAuthor> commentAuthor = ppt->GetCommentAuthors()->GetItem(i);
		for (int j = 0; j < commentAuthor->GetCommentsList()->GetCount(); j++)
		{
			//Get comment information
			intrusive_ptr<Comment> comment = commentAuthor->GetCommentsList()->GetItem(j);

			outFile << "Comment text : " << comment->GetText() << endl;
			outFile << "Comment author : " << comment->GetAuthorName() << endl;
			outFile << "Posted on time : " << comment->GetDateTime()->ToString() << endl;
		}
	}
	outFile.close();
}
