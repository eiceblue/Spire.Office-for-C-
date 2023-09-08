#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile =input_path+L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"GetDetailsOfSearchedText.txt";
	

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"Spire.PDF for .NET", Find_TextFindParameter::IgnoreCase);
	wstring str = L"";
	for(intrusive_ptr<PdfTextFind> find : collection->GetFinds())
	{
		str += L"==================================================================================";
		str += L"\nMatch Text: ";
		str += find->GetMatchText();
		str += L"\nText:";
		str += find->GetSearchText();
		str += L"\nSize:";
		str += find->GetSize()->ToString();
		str += L"\nPosition:";
		str += find->GetPosition()->ToString();
		str += L"\nThe index of page which is including the searched text : ";
		str += find->GetSearchPageIndex();
		str += L"\nThe line that contains the searched text : ";
		str += find->GetLineText();
		str += L"\nMatch Text: ";
		str += find->GetMatchText();
		str += L"\n";
	}
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << str;
	os.close();
	doc->Close();
}


