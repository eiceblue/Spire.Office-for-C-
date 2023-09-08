#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"RotateExistingPDF.pdf";
	wstring inputFile = DATAPATH"Sample.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page of the loaded PDF file
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Get the original rotation angle
	int rotation = static_cast<int>(page->GetRotation());

	//Set the angle
	rotation += static_cast<int>(PdfPageRotateAngle::RotateAngle270);

	//Rotate the PDF page based on
	page->SetRotation(static_cast<PdfPageRotateAngle>(rotation));
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
