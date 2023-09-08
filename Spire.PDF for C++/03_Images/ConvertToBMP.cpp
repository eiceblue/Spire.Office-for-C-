#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring fn1 = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertToBMP\\";
	
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile((fn1 + L"ToImage.pdf").c_str());
	//Save to images
	intrusive_ptr<Spire::Common::Stream> image = new Spire::Common::Stream();
	for (int i = 0; i < pdf->GetPages()->GetCount(); i++)
	{
		wstring s = to_wstring(i);
		wstring filename = outputFile + L"ToBMP-img-" + s + L".bmp";
		//var image
		image = pdf->SaveAsImage(i, Spire::Pdf::PdfImageType::Bitmap);
		image->Save(filename.c_str());
	}
	pdf->Close();
}









