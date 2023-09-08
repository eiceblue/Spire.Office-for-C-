#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Attachment.pdf";
	wstring input_Img = input_path + L"E-logo.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PdfAttachmentRelationship.pdf";

	//Load document from disk
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	//Define PdfAttachment
	intrusive_ptr<PdfAttachment> attachment = new PdfAttachment(input_Img.c_str());
	//Add addachment
	doc->GetAttachments()->Add(attachment, doc, PdfAttachmentRelationship::Alternative);
	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

