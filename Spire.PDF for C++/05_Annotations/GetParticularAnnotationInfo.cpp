#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_3.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetParticularAnnotationInfo.txt";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfAnnotationCollection> annotations = doc->GetPages()->GetItem(0)->GetAnnotationsWidget();
	wstring content;
	wchar_t nm_w[100];
	swprintf(nm_w, 100, L"%hs", typeid(PdfTextAnnotationWidget).name());
	LPCWSTR_S newName = nm_w;
	if (wcscmp(newName, annotations->GetItem(0)->GetInstanceTypeName()) == 0) {
		intrusive_ptr<PdfTextAnnotationWidget> textAnnotation = Object::Convert<PdfTextAnnotationWidget>(annotations->GetItem(0));
		content.append(L"Annotation text: ");
		content.append(textAnnotation->GetText());
		content.append(L"\nAnnotation ModifiedDate: ");
		content.append(textAnnotation->GetModifiedDate()->ToString());
		content.append(L"\nAnnotation author: ");
		content.append(textAnnotation->GetAuthor());
		content.append(L"\nAnnotation Name: ");
		content.append(textAnnotation->GetName());
	}
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << content;
	os.close();
	doc->Close();
}


