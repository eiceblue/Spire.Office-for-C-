#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_2.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile_p = output_path + L"attachment1.pdf";
	wstring outputFile_t = output_path + L"attachment1.txt";
	wstring outputFile_i = output_path + L"Logo.png";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get all attachments
	intrusive_ptr<PdfAttachmentCollection> attachments = doc->GetAttachments();

	attachments->GetItem(0)->GetData()->Save(outputFile_i.c_str());
	attachments->GetItem(1)->GetData()->Save(outputFile_p.c_str());
	attachments->GetItem(2)->GetData()->Save(outputFile_t.c_str());

	doc->Close();
}


