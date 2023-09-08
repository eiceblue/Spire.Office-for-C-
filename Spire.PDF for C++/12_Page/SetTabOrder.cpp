#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"SetTabOrder.pdf";
	wstring inputFile = DATAPATH"SetTabOrder.pdf";

	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());

	//Set using document structure
	pdf->GetFileInfo()->SetIncrementalUpdate(false);
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	page->SetTabOrder(TabOrder::Structure);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}
