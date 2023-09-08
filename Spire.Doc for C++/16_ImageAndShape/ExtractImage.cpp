#include "pch.h"
#include <deque>
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path;

	//open document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//document elements, each of them has child elements
	std::deque<intrusive_ptr<ICompositeObject>> nodes;
	nodes.push_back(document);

	//embedded images list.
#if defined(SKIASHARP)
	std::vector<SkiaSharp::SKintrusive_ptr<Image>> images;
#else
	std::vector<std::vector<byte>> images;
#endif
	//traverse
	while (nodes.size() > 0)
	{
		intrusive_ptr<ICompositeObject> node = nodes.front();
		nodes.pop_front();
		for (int i = 0; i < node->GetChildObjects()->GetCount(); i++)
		{
			intrusive_ptr<IDocumentObject> child = node->GetChildObjects()->GetItem(i);
			if (child->GetDocumentObjectType() == DocumentObjectType::Picture)
			{
				intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(child);
#if defined(SKIASHARP)
				SkiaSharp::SKintrusive_ptr<Image> image = SkiaSharp::SKImage::FromEncodedData(SkiaSharp::SKData::CreateCopy(picture->ImageBytes));
				images.push_back(image);
#else
				
				std::vector<byte> imageByte = picture->GetImageBytes();
				images.push_back(imageByte);
#endif
			}
			else if (Object::CheckType<ICompositeObject>(child))
			{
				nodes.push_back(boost::dynamic_pointer_cast<ICompositeObject>(child));
			}
		}
	}
#if defined(SKIASHARP)
	//save images         
	for (int i = 0; i < images.size(); i++)
	{
		std::wstring filename = StringHelper::formatSimple(outputFile.c_str() + L"Image-{0}.png", i);
		Fileintrusive_ptr<Stream> fileStream = new FileStream(filename, FileMode::Create, FileAccess::Write);
		images[i]->Encode(SkiaSharp::SKEncodedFREE_IMAGE_FORMAT::FIF_PNG, 100)->SaveTo(fileStream);
		fileStream->Flush();

		
	}
#else
	//save images
	for (int i = 0; i < images.size(); i++)
	{
		std::wstring fileName = L"Image-" + to_wstring(i) + L".png";
		std::ofstream outFile(outputFile + fileName, std::ios::binary);
		if (outFile.is_open())
		{
			outFile.write(reinterpret_cast<const char*>(images[i].data()), images[i].size());
			outFile.close();
		}
	}
#endif
	document->Close();


}
