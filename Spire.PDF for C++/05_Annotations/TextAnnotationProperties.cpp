#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"FreeTextAnnotation.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"TextAnnotationProperties.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfDocument> newDoc = new PdfDocument();
	for (int i = 0; i < page->GetAnnotationsWidget()->GetCount(); i++) {
		intrusive_ptr<PdfAnnotation> annotation = page->GetAnnotationsWidget()->GetItem(i);
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(PdfFreeTextAnnotationWidget).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, annotation->GetInstanceTypeName()) == 0) {
			intrusive_ptr<PdfFreeTextAnnotationWidget> textAnnotation = Object::Convert<PdfFreeTextAnnotationWidget>(annotation);
			intrusive_ptr<RectangleF> rect = textAnnotation->GetBounds();
			wstring text = textAnnotation->GetText();
			intrusive_ptr<PdfPageBase> newPage = newDoc->GetPages()->Add(page->GetSize());
			intrusive_ptr<PdfFreeTextAnnotation> newAnnotation = new PdfFreeTextAnnotation(rect);
			newAnnotation->SetText(text.c_str());
			newAnnotation->SetCalloutLines(textAnnotation->GetCalloutLines());
			newAnnotation->SetLineEndingStyle(textAnnotation->GetLineEndingStyle());
			newAnnotation->SetRectangleDifferences(textAnnotation->GetRectangularDifferenceArray());
			newAnnotation->SetColor(textAnnotation->GetColor());
			newAnnotation->SetAnnotationIntent(PdfAnnotationIntent::FreeTextCallout);
			newPage->GetAnnotationsWidget()->Add(newAnnotation);
		}
	}
	newDoc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
	newDoc->Close();
}

