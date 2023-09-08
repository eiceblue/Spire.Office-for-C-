#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;


const wstring ChildToString(PdfTextStyle style) {
	switch (style)
	{
	case Spire::Pdf::PdfTextStyle::Regular:
		return L"Regular";
		break;
	case Spire::Pdf::PdfTextStyle::Italic:
		return L"Italic";
		break;
	case Spire::Pdf::PdfTextStyle::Bold:
		return L"Bold";
		break;
	default:
		break;
	}
}
vector<wstring> SplitWstring(const wstring& text, const wstring& subText)
{
	vector<wstring> result;
	size_t left = 0;
	size_t right = text.find(subText);
	size_t textSize = text.size();
	size_t subTextSize = subText.size();

	while (right != wstring::npos)
	{
		if (right > left)
		{
			size_t size = right - left;
			result.push_back(text.substr(left, size));
			left = right + subTextSize;
		}
		else
			left += subTextSize;

		right = text.find(subText, left);
	}

	if (left < textSize)
		result.push_back(text.substr(left));

	return result;
}
void GetChildBookmarks(intrusive_ptr<PdfBookmarkCollection> bookmarks, std::wstring result)
{
	//Declare a new StringBuilder content
	wstring content;

	//Get Pdf bookmarks information.
	if (bookmarks->GetCount() > 0)
	{
		for (int i = 0; i < bookmarks->GetCount(); i++)
		{
			if (bookmarks->GetItem(i)->GetCount() > 0)
			{
				content += L"Child Bookmarks: \n";
				for (int j = 0; j < bookmarks->GetItem(i)->GetCount(); j++)
				{
					content += bookmarks->GetItem(i)->GetItem(j)->GetTitle();
					content += L"\n";
					wchar_t nm_w[100];
					swprintf(nm_w, 100, L"%hs", typeid(bookmarks->GetItem(i)->GetItem(j)->GetColor()).name());
					LPCWSTR_S typeName = nm_w;
					wstring tName = typeName;
					int bPos = tName.find(L"<");
					int ePos = tName.find(L">");
					tName = tName.substr(bPos + 1, ePos - bPos - 1);
					wstring subText = L" ";
					std::vector<wstring> seglist = SplitWstring(tName, subText);
					content += seglist[1] + L"\n";
					wstring textStyle = ChildToString(bookmarks->GetItem(i)->GetItem(j)->GetDisplayStyle());
					content += textStyle + L"\n";
				}
			}
		}
	}

	//Save to file.
	wofstream os;
	os.open(result, ios::trunc);
	os << content;
	os.close();
}

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetPdfChildBookmarks.txt";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	intrusive_ptr<PdfBookmarkCollection> bookmarks = doc->GetBookmarks();
	GetChildBookmarks(bookmarks, outputFile);

	doc->Close();
}
