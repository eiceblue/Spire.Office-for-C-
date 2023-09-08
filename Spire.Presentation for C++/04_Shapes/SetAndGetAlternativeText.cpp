#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ShapeTemplate.pptx";
	wstring outputFile = OUTPUTPATH"SetAlternativeText.pptx";
	wstring outputFile_txt = OUTPUTPATH"GetAlternativeText.txt";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Set the alternative text (title and description)
	slide->GetShapes()->GetItem(0)->SetAlternativeTitle(L"Rectangle");
	slide->GetShapes()->GetItem(0)->SetAlternativeText(L"This is a Rectangle");

	wstring* content = new wstring();

	//Get the alternative text (title and description)
	wstring title = slide->GetShapes()->GetItem(0)->GetAlternativeTitle();
	content->append(L"Title: " + title + L"\r\n");
	wstring description = slide->GetShapes()->GetItem(0)->GetAlternativeText();
	content->append(L"Description: " + description + L"\r\n");

	//Save the alternative text to Text file
	std::wofstream out;
	out.open(outputFile_txt);
	out.flush();
	out << content->c_str();
	out.close();

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

