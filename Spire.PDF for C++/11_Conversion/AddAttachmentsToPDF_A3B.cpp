#include "../pch.h"
using namespace Spire::Pdf;
using namespace std;

int main(){							
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddAttachmentsToPDF_A3B.pdf";

	PdfStandardsConverter* converter = new PdfStandardsConverter(inputFile.c_str());
	converter->ToPdfA3B(outputFile.c_str());
	delete converter;
}
