#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;


class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}

	StringBuilder* appendLine(const std::wstring& toAppend)
	{
		privateString += toAppend + L"\r\n";
		return this;
	}
	std::wstring toString()
	{
		return privateString;
	}
};

int main()
{
	wstring outputFile = OUTPUTPATH"GetDocumentProperties.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"PDFTemplate-Az.pdf";
	doc->LoadFromFile(inputFile.c_str());

	// Get document information
	intrusive_ptr<PdfDocumentInformation> docInfo = doc->GetDocumentInformation();

	// Create a StringBuilder object to put the details
	StringBuilder* builder = new StringBuilder();
	wstring result = L"";
	result += L"Author:";
	result += docInfo->GetAuthor();
	builder->appendLine(result);
	result = L"Creation Date: ";
	result += docInfo->GetCreationDate()->ToString();
	builder->appendLine(result);
	result = L"Keywords: ";
	result += docInfo->GetKeywords();
	builder->appendLine(result);
	result = L"Modify Date: ";
	result += docInfo->GetModificationDate()->ToString();
	builder->appendLine(result);
	result = L"Subject: ";
	result += docInfo->GetSubject();
	builder->appendLine(result);
	result = L"Title: ";
	result += docInfo->GetTitle();
	builder->appendLine(result);

	//Save to file.
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << builder->toString();
	os.close();

	doc->Close();
}