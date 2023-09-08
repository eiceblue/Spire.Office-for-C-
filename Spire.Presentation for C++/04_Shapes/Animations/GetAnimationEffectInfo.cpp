#include "pch.h"

using namespace Spire::Presentation;

namespace Animations
{
	const wstring effTypeToString(AnimationEffectType effType)
	{
		switch (effType)
		{
		case Spire::Presentation::AnimationEffectType::Appear:
			return L"Appear";
			break;
		case Spire::Presentation::AnimationEffectType::ArcUp:
			return L"ArcUp";
			break;
		case Spire::Presentation::AnimationEffectType::Ascend:
			return L"Ascend";
			break;
		case Spire::Presentation::AnimationEffectType::Blast:
			return L"Blast";
			break;
		case Spire::Presentation::AnimationEffectType::Blinds:
			return L"Blinds";
			break;
		case Spire::Presentation::AnimationEffectType::Blink:
			return L"Blink";
			break;
		case Spire::Presentation::AnimationEffectType::BoldFlash:
			return L"BoldFlash";
			break;
		case Spire::Presentation::AnimationEffectType::BoldReveal:
			return L"BoldReveal";
			break;
		case Spire::Presentation::AnimationEffectType::Boomerang:
			return L"Boomerang";
			break;
		case Spire::Presentation::AnimationEffectType::Bounce:
			return L"Bounce";
			break;
		case Spire::Presentation::AnimationEffectType::Box:
			return L"Box";
			break;
		case Spire::Presentation::AnimationEffectType::BrushOnColor:
			return L"BrushOnColor";
			break;
		case Spire::Presentation::AnimationEffectType::BrushOnUnderline:
			return L"BrushOnUnderline";
			break;
		case Spire::Presentation::AnimationEffectType::CenterRevolve:
			return L"CenterRevolve";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeFillColor:
			return L"ChangeFillColor";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeFont:
			return L"ChangeFont";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeFontColor:
			return L"ChangeFontColor";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeFontSize:
			return L"ChangeFontSize";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeFontStyle:
			return L"ChangeFontStyle";
			break;
		case Spire::Presentation::AnimationEffectType::ChangeLineColor:
			return L"ChangeLineColor";
			break;
		case Spire::Presentation::AnimationEffectType::Checkerboard:
			return L"Checkerboard";
			break;
		case Spire::Presentation::AnimationEffectType::Circle:
			return L"Circle";
			break;
		case Spire::Presentation::AnimationEffectType::ColorBlend:
			return L"ColorBlend";
			break;
		case Spire::Presentation::AnimationEffectType::ColorReveal:
			return L"ColorReveal";
			break;
		case Spire::Presentation::AnimationEffectType::ColorWave:
			return L"ColorWave";
			break;
		case Spire::Presentation::AnimationEffectType::ComplementaryColor:
			return L"ComplementaryColor";
			break;
		case Spire::Presentation::AnimationEffectType::ComplementaryColor2:
			return L"ComplementaryColor2";
			break;
		case Spire::Presentation::AnimationEffectType::Compress:
			return L"Compress";
			break;
		case Spire::Presentation::AnimationEffectType::ContrastingColor:
			return L"ContrastingColor";
			break;
		case Spire::Presentation::AnimationEffectType::Crawl:
			return L"Crawl";
			break;
		case Spire::Presentation::AnimationEffectType::Credits:
			return L"Credits";
			break;
		case Spire::Presentation::AnimationEffectType::Darken:
			return L"Darken";
			break;
		case Spire::Presentation::AnimationEffectType::Desaturate:
			return L"Desaturate";
			break;
		case Spire::Presentation::AnimationEffectType::Descend:
			return L"Descend";
			break;
		case Spire::Presentation::AnimationEffectType::Diamond:
			return L"Diamond";
			break;
		case Spire::Presentation::AnimationEffectType::Dissolve:
			return L"Dissolve";
			break;
		case Spire::Presentation::AnimationEffectType::EaseIn:
			return L"EaseIn";
			break;
		case Spire::Presentation::AnimationEffectType::Expand:
			return L"Expand";
			break;
		case Spire::Presentation::AnimationEffectType::Fade:
			return L"Fade";
			break;
		case Spire::Presentation::AnimationEffectType::FadedSwivel:
			return L"FadedSwivel";
			break;
		case Spire::Presentation::AnimationEffectType::FadedZoom:
			return L"FadedZoom";
			break;
		case Spire::Presentation::AnimationEffectType::FlashBulb:
			return L"FlashBulb";
			break;
		case Spire::Presentation::AnimationEffectType::FlashOnce:
			return L"FlashOnce";
			break;
		case Spire::Presentation::AnimationEffectType::Flicker:
			return L"Flicker";
			break;
		case Spire::Presentation::AnimationEffectType::Flip:
			return L"Flip";
			break;
		case Spire::Presentation::AnimationEffectType::Float:
			return L"Float";
			break;
		case Spire::Presentation::AnimationEffectType::Fly:
			return L"Fly";
			break;
		case Spire::Presentation::AnimationEffectType::Fold:
			return L"Fold";
			break;
		case Spire::Presentation::AnimationEffectType::Glide:
			return L"Glide";
			break;
		case Spire::Presentation::AnimationEffectType::GrowAndTurn:
			return L"GrowAndTurn";
			break;
		case Spire::Presentation::AnimationEffectType::GrowShrink:
			return L"GrowShrink";
			break;
		case Spire::Presentation::AnimationEffectType::GrowWithColor:
			return L"GrowWithColor";
			break;
		case Spire::Presentation::AnimationEffectType::Lighten:
			return L"Lighten";
			break;
		case Spire::Presentation::AnimationEffectType::LightSpeed:
			return L"LightSpeed";
			break;
		case Spire::Presentation::AnimationEffectType::MediaPause:
			return L"MediaPause";
			break;
		case Spire::Presentation::AnimationEffectType::MediaPlay:
			return L"MediaPlay";
			break;
		case Spire::Presentation::AnimationEffectType::MediaStop:
			return L"MediaStop";
			break;
		case Spire::Presentation::AnimationEffectType::Path4PointStar:
			return L"Path4PointStar";
			break;
		case Spire::Presentation::AnimationEffectType::Path5PointStar:
			return L"Path5PointStar";
			break;
		case Spire::Presentation::AnimationEffectType::Path6PointStar:
			return L"Path6PointStar";
			break;
		case Spire::Presentation::AnimationEffectType::Path8PointStar:
			return L"Path8PointStar";
			break;
		case Spire::Presentation::AnimationEffectType::PathArcDown:
			return L"PathArcDown";
			break;
		case Spire::Presentation::AnimationEffectType::PathArcLeft:
			return L"PathArcLeft";
			break;
		case Spire::Presentation::AnimationEffectType::PathArcRight:
			return L"PathArcRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathArcUp:
			return L"PathArcUp";
			break;
		case Spire::Presentation::AnimationEffectType::PathBean:
			return L"PathBean";
			break;
		case Spire::Presentation::AnimationEffectType::PathBounceLeft:
			return L"PathBounceLeft";
			break;
		case Spire::Presentation::AnimationEffectType::PathBounceRight:
			return L"PathBounceRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathBuzzsaw:
			return L"PathBuzzsaw";
			break;
		case Spire::Presentation::AnimationEffectType::PathCircle:
			return L"PathCircle";
			break;
		case Spire::Presentation::AnimationEffectType::PathCrescentMoon:
			return L"PathCrescentMoon";
			break;
		case Spire::Presentation::AnimationEffectType::PathCurvedSquare:
			return L"PathCurvedSquare";
			break;
		case Spire::Presentation::AnimationEffectType::PathCurvedX:
			return L"PathCurvedX";
			break;
		case Spire::Presentation::AnimationEffectType::PathCurvyLeft:
			return L"PathCurvyLeft";
			break;
		case Spire::Presentation::AnimationEffectType::PathCurvyRight:
			return L"PathCurvyRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathCurvyStar:
			return L"PathCurvyStar";
			break;
		case Spire::Presentation::AnimationEffectType::PathDecayingWave:
			return L"PathDecayingWave";
			break;
		case Spire::Presentation::AnimationEffectType::PathDiagonalDownRight:
			return L"PathDiagonalDownRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathDiagonalUpRight:
			return L"PathDiagonalUpRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathDiamond:
			return L"PathDiamond";
			break;
		case Spire::Presentation::AnimationEffectType::PathDown:
			return L"PathDown";
			break;
		case Spire::Presentation::AnimationEffectType::PathEqualTriangle:
			return L"PathEqualTriangle";
			break;
		case Spire::Presentation::AnimationEffectType::PathFigure8Four:
			return L"PathFigure8Four";
			break;
		case Spire::Presentation::AnimationEffectType::PathFootball:
			return L"PathFootball";
			break;
		case Spire::Presentation::AnimationEffectType::PathFunnel:
			return L"PathFunnel";
			break;
		case Spire::Presentation::AnimationEffectType::PathHeart:
			return L"PathHeart";
			break;
		case Spire::Presentation::AnimationEffectType::PathHeartbeat:
			return L"PathHeartbeat";
			break;
		case Spire::Presentation::AnimationEffectType::PathHexagon:
			return L"PathHexagon";
			break;
		case Spire::Presentation::AnimationEffectType::PathHorizontalFigure8:
			return L"PathHorizontalFigure8";
			break;
		case Spire::Presentation::AnimationEffectType::PathInvertedSquare:
			return L"PathInvertedSquare";
			break;
		case Spire::Presentation::AnimationEffectType::PathInvertedTriangle:
			return L"PathInvertedTriangle";
			break;
		case Spire::Presentation::AnimationEffectType::PathLeft:
			return L"PathLeft";
			break;
		case Spire::Presentation::AnimationEffectType::PathLoopdeLoop:
			return L"PathLoopdeLoop";
			break;
		case Spire::Presentation::AnimationEffectType::PathNeutron:
			return L"PathNeutron";
			break;
		case Spire::Presentation::AnimationEffectType::PathOctagon:
			return L"PathOctagon";
			break;
		case Spire::Presentation::AnimationEffectType::PathParallelogram:
			return L"PathParallelogram";
			break;
		case Spire::Presentation::AnimationEffectType::PathPeanut:
			return L"PathPeanut";
			break;
		case Spire::Presentation::AnimationEffectType::PathPentagon:
			return L"PathPentagon";
			break;
		case Spire::Presentation::AnimationEffectType::PathPlus:
			return L"PathPlus";
			break;
		case Spire::Presentation::AnimationEffectType::PathPointyStar:
			return L"PathPointyStar";
			break;
		case Spire::Presentation::AnimationEffectType::PathRight:
			return L"PathRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathRightTriangle:
			return L"PathRightTriangle";
			break;
		case Spire::Presentation::AnimationEffectType::PathSCurve1:
			return L"PathSCurve1";
			break;
		case Spire::Presentation::AnimationEffectType::PathSCurve2:
			return L"PathSCurve2";
			break;
		case Spire::Presentation::AnimationEffectType::PathSineWave:
			return L"PathSineWave";
			break;
		case Spire::Presentation::AnimationEffectType::PathSpiralLeft:
			return L"PathSpiralLeft";
			break;
		case Spire::Presentation::AnimationEffectType::PathSpiralRight:
			return L"PathSpiralRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathSpring:
			return L"PathSpring";
			break;
		case Spire::Presentation::AnimationEffectType::PathSquare:
			return L"PathSquare";
			break;
		case Spire::Presentation::AnimationEffectType::PathStairsDown:
			return L"PathStairsDown";
			break;
		case Spire::Presentation::AnimationEffectType::PathSwoosh:
			return L"PathSwoosh";
			break;
		case Spire::Presentation::AnimationEffectType::PathTeardrop:
			return L"PathTeardrop";
			break;
		case Spire::Presentation::AnimationEffectType::PathTrapezoid:
			return L"PathTrapezoid";
			break;
		case Spire::Presentation::AnimationEffectType::PathTurnDown:
			return L"PathTurnDown";
			break;
		case Spire::Presentation::AnimationEffectType::PathTurnRight:
			return L"PathTurnRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathTurnUp:
			return L"PathTurnUp";
			break;
		case Spire::Presentation::AnimationEffectType::PathTurnUpRight:
			return L"PathTurnUpRight";
			break;
		case Spire::Presentation::AnimationEffectType::PathUp:
			return L"PathUp";
			break;
		case Spire::Presentation::AnimationEffectType::PathUser:
			return L"PathUser";
			break;
		case Spire::Presentation::AnimationEffectType::PathVerticalFigure8:
			return L"PathVerticalFigure8";
			break;
		case Spire::Presentation::AnimationEffectType::PathWave:
			return L"PathWave";
			break;
		case Spire::Presentation::AnimationEffectType::PathZigzag:
			return L"PathZigzag";
			break;
		case Spire::Presentation::AnimationEffectType::Peek:
			return L"Peek";
			break;
		case Spire::Presentation::AnimationEffectType::Pinwheel:
			return L"Pinwheel";
			break;
		case Spire::Presentation::AnimationEffectType::Plus:
			return L"Plus";
			break;
		case Spire::Presentation::AnimationEffectType::RandomBars:
			return L"RandomBars";
			break;
		case Spire::Presentation::AnimationEffectType::RandomEffects:
			return L"RandomEffects";
			break;
		case Spire::Presentation::AnimationEffectType::RiseUp:
			return L"RiseUp";
			break;
		case Spire::Presentation::AnimationEffectType::Shimmer:
			return L"Shimmer";
			break;
		case Spire::Presentation::AnimationEffectType::Sling:
			return L"Sling";
			break;
		case Spire::Presentation::AnimationEffectType::Spin:
			return L"Spin";
			break;
		case Spire::Presentation::AnimationEffectType::Spinner:
			return L"Spinner";
			break;
		case Spire::Presentation::AnimationEffectType::Spiral:
			return L"Spiral";
			break;
		case Spire::Presentation::AnimationEffectType::Split:
			return L"Split";
			break;
		case Spire::Presentation::AnimationEffectType::Stretch:
			return L"Stretch";
			break;
		case Spire::Presentation::AnimationEffectType::Strips:
			return L"Strips";
			break;
		case Spire::Presentation::AnimationEffectType::StyleEmphasis:
			return L"StyleEmphasis";
			break;
		case Spire::Presentation::AnimationEffectType::Swish:
			return L"Swish";
			break;
		case Spire::Presentation::AnimationEffectType::Swivel:
			return L"Swivel";
			break;
		case Spire::Presentation::AnimationEffectType::Teeter:
			return L"Teeter";
			break;
		case Spire::Presentation::AnimationEffectType::Thread:
			return L"Thread";
			break;
		case Spire::Presentation::AnimationEffectType::Transparency:
			return L"Transparency";
			break;
		case Spire::Presentation::AnimationEffectType::Unfold:
			return L"Unfold";
			break;
		case Spire::Presentation::AnimationEffectType::VerticalGrow:
			return L"VerticalGrow";
			break;
		case Spire::Presentation::AnimationEffectType::Wave:
			return L"Wave";
			break;
		case Spire::Presentation::AnimationEffectType::Wedge:
			return L"Wedge";
			break;
		case Spire::Presentation::AnimationEffectType::Wheel:
			return L"Wheel";
			break;
		case Spire::Presentation::AnimationEffectType::Whip:
			return L"Whip";
			break;
		case Spire::Presentation::AnimationEffectType::Wipe:
			return L"Wipe";
			break;
		case Spire::Presentation::AnimationEffectType::Magnify:
			return L"Magnify";
			break;
		case Spire::Presentation::AnimationEffectType::Zoom:
			return L"Zoom";
			break;
		default:
			break;
		}
	}
}
int main()
{

	wstring inputFile = DATAPATH"Animation.pptx";
	wstring outputFile = OUTPUTPATH"GetAnimationEffectInfo.txt";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	wofstream outFile(outputFile, ios::out);
	//Travel each slide

	for (int l = 0; l < presentation->GetSlides()->GetCount(); l++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(l);
		for (int e = 0; e < slide->GetTimeline()->GetMainSequence()->GetCount(); e++)
		{
			intrusive_ptr<AnimationEffect> effect = slide->GetTimeline()->GetMainSequence()->GetItem(e);
			//Get the animation effect type
			AnimationEffectType animationEffectType = effect->GetAnimationEffectType();
			outFile << "animation effect type:" << Animations::effTypeToString(animationEffectType) << endl;
			//Get the slide number where the animation is located
			int slideNumber = slide->GetSlideNumber();
			outFile << "slide number:" << std::to_string(slideNumber).c_str() << endl;
			//Get the shape name
			wstring shapeName = effect->GetShapeTarget()->GetName();
			outFile << "shape name:" << shapeName.c_str() << endl;
		}
	}
	outFile.close();
}

