#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring input_1 = input_path + L"SampleB_1.pdf";
	wstring input_2 = input_path + L"SampleB_2.pdf";
	wstring input_3 = input_path + L"SampleB_3.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SortFileInPdf.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();

	doc->GetCollection()->AddCustomField(L"No", L"number", CustomFieldType::NumberField);
	doc->GetCollection()->AddFileRelatedField(L"Desc", L"desc", FileRelatedFieldType::Desc);
	doc->GetCollection()->Sort({ L"No", L"Desc" }, { true, true });

	intrusive_ptr<PdfAttachment> pdfAttachment = new PdfAttachment(input_1.c_str());
	doc->GetCollection()->AddAttachment(pdfAttachment);
	pdfAttachment = new PdfAttachment(input_2.c_str());
	doc->GetCollection()->AddAttachment(pdfAttachment);
	pdfAttachment = new PdfAttachment(input_3.c_str());
	doc->GetCollection()->AddAttachment(pdfAttachment);

	int i = 1;
	std::vector<intrusive_ptr<PdfAttachment>> attachment = doc->GetCollection()->GetAssociatedFiles();
	for (int j = 0; j < attachment.size(); j++)
	{
		attachment[j]->SetFieldValue(L"No", j + 1);
		attachment[j]->SetFieldValue(L"Desc", attachment[j]->GetFileName());
	}
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

