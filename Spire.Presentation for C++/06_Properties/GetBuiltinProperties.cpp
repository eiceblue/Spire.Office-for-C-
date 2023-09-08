#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"GetProperties.pptx";
	wstring outputFile = OUTPUTPATH"GetBuiltinProperties.txt";

	//Create PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the PPT document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the builtin properties 
	wstring application = presentation->GetDocumentProperty()->GetApplication();
	wstring author = presentation->GetDocumentProperty()->GetAuthor();
	wstring company = presentation->GetDocumentProperty()->GetCompany();
	wstring keywords = presentation->GetDocumentProperty()->GetKeywords();
	wstring comments = presentation->GetDocumentProperty()->GetComments();
	wstring category = presentation->GetDocumentProperty()->GetCategory();
	wstring title = presentation->GetDocumentProperty()->GetTitle();
	wstring subject = presentation->GetDocumentProperty()->GetSubject();

	wofstream outFile(outputFile, ios::out);
	outFile << "DocumentProperty.Application: " << application.c_str() << endl;
	outFile << "DocumentProperty.Author: " << author.c_str() << endl;
	outFile << "DocumentProperty.Company " << company.c_str() << endl;
	outFile << "DocumentProperty.Keywords: " << keywords.c_str() << endl;
	outFile << "DocumentProperty.Comments: " << comments.c_str() << endl;
	outFile << "DocumentProperty.Category: " << category.c_str() << endl;
	outFile << "DocumentProperty.Title: " << title.c_str() << endl;
	outFile << "DocumentProperty.Subject: " << subject.c_str();

	//Save them to a txt file
	outFile.close();
}

