#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_2.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetPdfAttachmentInfo.txt";

	//Create a new PDF document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();

	//Load the file from disk.
	pdf->LoadFromFile(inputFile.c_str());

	//Get a collection of attachments on the PDF document
	intrusive_ptr<PdfAttachmentCollection> collection = pdf->GetAttachments();

	//Get the first attachment.
	intrusive_ptr<PdfAttachment> attachment = collection->GetItem(0);

	//Get the information of the first attachment.
	wstring content = L"";
	content += L"Filename: " + (wstring)attachment->GetFileName() + L"\r\n";
	content += L"Description: " + (wstring)attachment->GetDescription() + L"\r\n";
	content += L"Creation Date: " + (wstring)attachment->GetCreationDate()->ToString() + L"\r\n";
	content += L"Modification Date: " + (wstring)attachment->GetModificationDate()->ToString();
	//Save to file.
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << content;
	os.close();

	pdf->Close();
}


