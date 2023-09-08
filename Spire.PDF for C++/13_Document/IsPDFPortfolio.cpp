#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"TextBoxSampleB_1.pdf";
	wstring outputFile = OUTPUTPATH"IsPDFPortfolio.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->LoadFromFile(inputFile.c_str());

	//Judge whether the document is portfolio or not.
	bool value = doc->GetIsPortfolio();

	if (value)
	{
		//Save to file.
		wofstream os;
		os.open(outputFile, ios::trunc);
		os << L"The document is portfolio";
		os.close();
	}
	else
	{
		//Save to file.
		wofstream os;
		os.open(outputFile, ios::trunc);
		os << L"The document is not portfolio";
		os.close();
	}

	doc->Close();
}
