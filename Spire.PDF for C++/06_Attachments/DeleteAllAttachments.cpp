#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DeleteAllAttachments.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DeleteAllAttachments.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get all attachments
	intrusive_ptr<PdfAttachmentCollection> attachments = doc->GetAttachments();

	//Delete all attachments
	attachments->Clear();
	//Save pdf file
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}


