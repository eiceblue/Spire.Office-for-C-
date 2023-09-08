#include "pch.h"


using namespace Spire::Pdf;
using namespace std;

void AddAttachmentsToPDF_X1A2001()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_X1A2001.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfX1A2001(outputFile.c_str());
	converter->Dispose();
}

void AddAttachmentsToPDF_A1A()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A1A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA1A(outputFile.c_str());
}

void AddAttachmentsToPDF_A2A()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A2A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA2A(outputFile.c_str());
}

void AddAttachmentsToPDF_A2B()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A2B.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA2B(outputFile.c_str());
}

void AddAttachmentsToPDF_A3A()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A3A.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA3A(outputFile.c_str());
}

void AddAttachmentsToPDF_A3B()
{
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A3B.pdf";
	wstring inputFile = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA3B(outputFile.c_str());
}

int main()
{
	wstring inputFile = DATAPATH"SampleB_2.pdf";
	wstring outputFile = OUTPUTPATH"AddAttachmentsToPDF_A1B.pdf";
	wstring output_ = OUTPUTPATH"Temp/AddAttachmentsToPDF_Temp.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfNewDocument> newDoc = new PdfNewDocument();
	newDoc->SetConformance(PdfConformanceLevel::Pdf_A1B);
	for (int i = 0; i < doc->GetPages()->GetCount(); i++) {
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		intrusive_ptr<SizeF> size = page->GetSize();
		intrusive_ptr<PdfPageBase> p = newDoc->GetPages()->Add(size, new PdfMargins(0));
		page->CreateTemplate()->Draw(p, 0.f, 0.f);
	}

	//Load files and add in attachments
	wstring data1 = DATAPATH"SampleB_1.png";
	wstring data2 = DATAPATH"SampleB_1.pdf";
	intrusive_ptr<Stream> s1 = new Stream(data1.c_str());
	intrusive_ptr<PdfAttachment> attach1 = new PdfAttachment(L"attachment1.png", s1);
	intrusive_ptr<Stream> s2 = new Stream(data2.c_str());
	intrusive_ptr<PdfAttachment> attach2 = new PdfAttachment(L"attachment2.pdf", s2);
	newDoc->GetAttachments()->Add(attach1);
	newDoc->GetAttachments()->Add(attach2);
	intrusive_ptr<Stream> file = new Stream(output_.c_str());
	newDoc->Save(file);
	newDoc->Close(true);
	intrusive_ptr<PdfStandardsConverter> converter = new PdfStandardsConverter(file);
	converter->ToPdfA1B(outputFile.c_str());

	//AddAttachmentsToPDF_X1A2001();
	//AddAttachmentsToPDF_A1A();
	//AddAttachmentsToPDF_A2A();
	//AddAttachmentsToPDF_A2B();
	//AddAttachmentsToPDF_A3A();
	//AddAttachmentsToPDF_A3B();
}
