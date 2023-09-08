#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"DeleteComment.pptx";
	wstring outputFile = OUTPUTPATH"DeleteComment.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Replace the text in the comment
	ppt->GetSlides()->GetItem(0)->GetComments()[1]->SetText(L"Replace comment");

	////delete the third comment
	ppt->GetSlides()->GetItem(0)->DeleteComment(ppt->GetSlides()->GetItem(0)->GetComments()[2]);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
