#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}

	StringBuilder* appendLine(const std::wstring& toAppend)
	{
		privateString += toAppend + L"\r\n";
		return this;
	}

	std::wstring toString()
	{
		return privateString;
	}
};
int main()
{
	wstring inputFile = DATAPATH"LinkAnnotation.pdf";
	wstring outputFile = OUTPUTPATH"GetLinkAnnotation.txt";

	//Create a pdf document
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	//Load file from disk
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//Get the annotation collection
	intrusive_ptr<PdfAnnotationCollection> annotations = page->GetAnnotationsWidget();

	//Create StringBuilder to save 
	StringBuilder* content = new StringBuilder();

	//Verify whether widgetCollection is not null or not
	if (annotations->GetCount() > 0)
	{
		intrusive_ptr<PdfTextWebLinkAnnotationWidget> WebLinkAnnotation;
		for (int i = 0; i < page->GetAnnotationsWidget()->GetCount(); i++) {
			wchar_t nm_w[100];
			swprintf(nm_w, 100, L"%hs", typeid(PdfTextWebLinkAnnotationWidget).name());
			LPCWSTR_S newName = nm_w;
			if (wcscmp(newName, page->GetAnnotationsWidget()->GetItem(i)->GetInstanceTypeName()) == 0)
			{
				WebLinkAnnotation = Object::Convert<PdfTextWebLinkAnnotationWidget>(page->GetAnnotationsWidget()->GetItem(i));
				std::wstring url = WebLinkAnnotation->GetUrl();
				content->appendLine(L"The url of link annotation is " + url);
				std::wstring strC = L"The text of link annotation is ";
				strC += WebLinkAnnotation->GetText();
				content->appendLine(strC);
			}
		}
	}
	wstring str = content->toString();
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << str;
	os.close();
	doc->Close();
}

