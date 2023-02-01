#include "pch.h"
#include <deque>
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractImage/";

	//open document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//document elements, each of them has child elements
	deque<ICompositeObject*> nodes;
	nodes.push_back(document);

	//embedded images list.
	vector<Image*> images;
	//traverse
	while (nodes.size() > 0)
	{
		ICompositeObject* node = nodes.front();
		nodes.pop_front();
		for (int  i =0;i<node->GetChildObjects()->GetCount();i++)
		{
			IDocumentObject* child = node->GetChildObjects()->GetItem(i);
			if (child->GetDocumentObjectType() == DocumentObjectType::Picture)
			{
				DocPicture* picture = dynamic_cast<DocPicture*>(child);
				images.push_back(picture->GetImage());
			}
			else if (dynamic_cast<ICompositeObject*>(child) != nullptr)
			{
				nodes.push_back(dynamic_cast<ICompositeObject*>(child));
			}

		}
	}
	//save images
	for (int i = 0; i < images.size(); i++)
	{
		wstring fileName = L"Image-" + to_wstring(i) + L".png";
		images[i]->Save((outputFile + fileName).c_str(), ImageFormat::GetPng());
	}
	document->Close();
	delete document;
}
