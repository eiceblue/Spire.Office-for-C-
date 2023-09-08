#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_5.pptx";
	wstring outputFile = OUTPUTPATH"ExtractComments.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);

	//Get all comments from the first slide.
	std::vector<intrusive_ptr<Comment>> comments = ppt->GetSlides()->GetItem(0)->GetComments();

	//Save the comments in txt file.
	for (int i = 0; i < comments.size(); i++)
	{
		outFile << comments[i]->GetText() << endl;
	}
	outFile.close();
}
