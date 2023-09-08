#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"UpdateFreeTextAnnotation.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateFreeTextAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfAnnotationCollection> annotations = doc->GetPages()->GetItem(0)->GetAnnotationsWidget();
	for (int i = 0; i < annotations->GetCount(); i++) {
		intrusive_ptr<PdfFreeTextAnnotationWidget> annotaion = Object::Convert<PdfFreeTextAnnotationWidget>(annotations->GetItem(i));
		annotaion->SetColor(new PdfRGBColor(Spire::Common::Color::GetYellowGreen()));
	}

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


