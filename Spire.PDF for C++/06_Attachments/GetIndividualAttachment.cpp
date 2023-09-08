#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_2.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetIndividualAttachment.pdf";

	//Create a new PDF document.
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();

	//Load the file from disk.
	pdf->LoadFromFile(inputFile.c_str());

	//Get a collection of attachments on the PDF document.
	intrusive_ptr<PdfAttachmentCollection> collection = pdf->GetAttachments();

	//Get the second attachment in PDF file.
	collection->GetItem(1)->GetData()->Save(outputFile.c_str());

	pdf->Close();
}

