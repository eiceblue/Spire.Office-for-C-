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
	wstring outputFile = OUTPUTPATH"GetPageLabel.txt";
	wstring inputFile = DATAPATH"PageLabel.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	//Create a StringBuilder instance
	StringBuilder* sb = new StringBuilder();

	//Get the lables of the pages in the PDF file
	for (int i = 0; i < pdf->GetPages()->GetCount(); i++)
	{
		sb->appendLine(L"The page lable of page " + std::to_wstring(i + 1) + L" is \"" + pdf->GetPages()->GetItem(i)->GetPageLabel() + L"\"");
	}

	wofstream os;
	os.open(outputFile, ios::trunc);
	os << sb->toString();
	os.close();
}
