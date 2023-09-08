#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile_pptx = OUTPUTPATH"SetPropertiesForTemplate.pptx";
	wstring outputFile_ppt = OUTPUTPATH"SetPropertiesForTemplate.ppt";
	wstring outputFile_odp = OUTPUTPATH"SetPropertiesForTemplate.odp";  //保存ODP 异常

	//Create a document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Set the DocumentProperty 
	presentation->GetDocumentProperty()->SetApplication(L"Spire.Presentation");
	presentation->GetDocumentProperty()->SetAuthor(L"E-iceblue");
	presentation->GetDocumentProperty()->SetCompany(L"E-iceblue Co., Ltd.");
	presentation->GetDocumentProperty()->SetKeywords(L"Demo File");
	presentation->GetDocumentProperty()->SetComments(L"This file is used to test Spire.Presentation.");
	presentation->GetDocumentProperty()->SetCategory(L"Demo");
	presentation->GetDocumentProperty()->SetTitle(L"This is a demo file.");
	presentation->GetDocumentProperty()->SetSubject(L"Test");

	//Create the .pptx template
	presentation->SaveToFile(outputFile_pptx.c_str(), FileFormat::Pptx2013);

	//Create the .odp template
	presentation->SaveToFile(outputFile_odp.c_str(), FileFormat::ODP); //保存ODP 异常

	//Create the .ppt template
	presentation->SaveToFile(outputFile_ppt.c_str(), FileFormat::PPT);
}

