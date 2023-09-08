#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;


int main()
{
	wstring inputFile = DATAPATH"SplitDocument.pdf";
	wstring outputFile = OUTPUTPATH"SplitDocument/";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	std::wstring pattern = outputFile + L"SplitDocument-{0}.pdf";
	//Split document
	doc->Split(pattern.c_str());

	doc->Close();
	
}
