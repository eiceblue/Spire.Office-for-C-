#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

static std::wstring trimStart(std::wstring source, const std::wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(0, source.find_first_not_of(trimChars));
}

static std::wstring trimEnd(std::wstring source, const std::wstring& trimChars = L" \t\n\r\v\f")
{
	return source.erase(source.find_last_not_of(trimChars) + 1);
}

static std::wstring trim(std::wstring source, const std::wstring& trimChars = L" \t\n\r\v\f")
{
	return trimStart(trimEnd(source, trimChars), trimChars);
}

template<typename T>
static std::wstring formatSimple(const std::wstring& input, T arg)
{
	std::wostringstream ss;
	std::size_t lastCloseBrace = std::wstring::npos;
	std::size_t openBrace = std::wstring::npos;
	while ((openBrace = input.find(L'{', openBrace + 1)) != std::wstring::npos)
	{
		if (openBrace + 1 < input.length())
		{
			if (input[openBrace + 1] == L'{')
			{
				openBrace++;
				continue;
			}

			std::size_t closeBrace = input.find(L'}', openBrace + 1);
			if (closeBrace != std::wstring::npos)
			{
				ss << input.substr(lastCloseBrace + 1, openBrace - lastCloseBrace - 1);
				lastCloseBrace = closeBrace;

				std::wstring index = trim(input.substr(openBrace + 1, closeBrace - openBrace - 1));
				if (index == L"0")
					ss << arg;
				else
					throw std::runtime_error("Only simple positional format specifiers are handled by the 'formatSimple' helper method.");
			}
		}
	}

	if (lastCloseBrace + 1 < input.length())
		ss << input.substr(lastCloseBrace + 1);

	return ss.str();
}

class StringBuilder
{
private:
	std::wstring privateString;

public:
	StringBuilder()
	{
	}
	StringBuilder* append(const std::wstring& toAppend)
	{
		privateString += toAppend;
		return this;
	}
	std::wstring toString()
	{
		return privateString;
	}
};

int main()
{
	wstring outputFile = OUTPUTPATH"Extraction.txt";
	wstring outputFile_I = OUTPUTPATH"Extraction/";
	wstring inputFile = DATAPATH"Extraction.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	StringBuilder* buffer = new StringBuilder();
	
	std::vector<intrusive_ptr<Image>> images;
	for (int i = 0; i < doc->GetPages()->GetCount(); i++)
	{
		intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(i);
		buffer->append(page->ExtractText());
		// foreach (SkiaSharp.SKImage image in page.ExtractImages())
		std::vector<intrusive_ptr<Image>> exImages = page->ExtractImages();
		for (int j = 0; j < exImages.size(); j++)
		{
			images.push_back(exImages[j]);
		}
	}

	doc->Close();
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << buffer->toString();
	os.close();

	//save image
	int index = 0;
	for (auto image : images)
	{
		std::wstring imageFileName = formatSimple(outputFile_I + L"Image-{0}.png", index++);
		image->Save(imageFileName.c_str(), ImageFormat::GetPng());
	}
}
