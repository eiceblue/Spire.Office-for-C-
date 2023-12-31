#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"Properties.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	wstring inputFile = DATAPATH"Properties.pdf";
	doc->LoadFromFile(inputFile.c_str());

	//Set document info
	doc->GetDocumentInformation()->SetAuthor(L"E-iceblue");
	doc->GetDocumentInformation()->SetCreator(L"E-iceblue");
	doc->GetDocumentInformation()->SetKeywords(L"pdf, demo, document information");
	doc->GetDocumentInformation()->SetProducer(L"Spire.Pdf");
	doc->GetDocumentInformation()->SetSubject(L"Demo of Spire.Pdf");
	doc->GetDocumentInformation()->SetTitle(L"Document Information");

	//File info
	doc->GetFileInfo()->SetCrossReferenceType(PdfCrossReferenceType::CrossReferenceStream);
	doc->GetFileInfo()->SetIncrementalUpdate(false);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
