#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"MergePdfsByStream.pdf";	
	wstring input1 = DATAPATH"MergePdfsTemplate_1.pdf";
	wstring input2 = DATAPATH"MergePdfsTemplate_2.pdf";
	wstring input3 = DATAPATH"MergePdfsTemplate_3.pdf";
	intrusive_ptr<Stream> stream1 = new Stream(input1.c_str());
	intrusive_ptr<Stream> stream2 = new Stream(input2.c_str());
	intrusive_ptr<Stream> stream3 = new Stream(input3.c_str());
	//Pdf document streams 
	std::vector< intrusive_ptr<Stream>> streams = { stream1, stream2, stream3 };

	//Also can merge files by filename
	//Merge files by stream
	intrusive_ptr<PdfDocumentBase> doc = PdfDocument::MergeFiles(streams);

	//Save the document
	doc->Save(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}
