/*
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2008 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------
*/



// class ClideBox
function CSlideBox()
{
	this.m_oDiv = null;
	
	this.setBox = function( o )
	{
		this.m_oDiv = o;
	};
	
	this.getBox = function()
	{
		return this.m_oDiv;
	};


	this.height = function()
	{
		var	o = this.m_oDiv;
		var h;
		
		if ( o.offsetHeight )
			h = o.offsetHeight;
		else if ( o.style.pixelHeight )
			h = o.style.pixelHeight;
		else if ( o.height )
			h = o.height;
	
		return h;
	};




}


CSlideBox.prototype.setHeight
= function( nHeight )
{
	this.m_oDiv.style.height = nHeight+"px";
};


CSlideBox.prototype.setTop
= function( nTop )
{
	this.m_oDiv.style.top = nTop+"px";
};


CSlideBox.prototype.width
= function()
{
	var	o = this.m_oDiv;
	var w;
	if ( o.offsetWidth )
		w = o.offsetWidth;
	else if ( o.style.pixelWidth )
		w = o.style.pixelWidth;
	else if ( o.width )
		w = o.width;
	return w;
};


CSlideBox.prototype.setWidth
= function( nWidth )
{
	this.m_oDiv.style.width = nWidth+"px";
};





CSlideBox.prototype.loadContent
= function( sHTML )
{
	if ( document.layers )
	{
		this.m_oDiv.document.write(sHTML);
		this.m_oDiv.document.close();
	}
	else if ( document.all || document.getElementById )
	{
		//window.alert( "div="+ this.m_oDiv.id +"\n loadContent = " + sHTML );
		this.m_oDiv.innerHTML = sHTML;
	}
};









// class CSlideScreen : CSlideBox
//		encapsulates the box that represents the screen

function CSlideScreen( sDivID )
{
	CSlideBox.call( this );
	this.m_sDivID = sDivID;
}

CSlideScreen.prototype = new CSlideBox;
CSlideScreen.prototype.constructor = CSlideScreen;



CSlideScreen.prototype.initialize
= function()
{
	var o = document.getElementById( this.m_sDivID );
	if ( o )
	{
		this.setBox( o );
		return true;
	}
	else
	{
		return false;
	}
};


CSlideScreen.prototype.getCanvas
= function()
{
	return this.getBox();
};



CSlideScreen.prototype.displaySlide
= function( sSlideValue )
{
	this.displaySlidePreprocess( sSlideValue );
	this.loadContent( sSlideValue );
	this.displaySlidePostprocess( sSlideValue );
};


CSlideScreen.prototype.displaySlidePreprocess
= function( sSlideValue )
{
};

CSlideScreen.prototype.displaySlidePostprocess
= function( sSlideValue )
{
};





//	class CSlideScreenTransition : CSlideScreen

function CSlideScreenTransition( sDivID )
{
	CSlideScreen.call( this, sDivID );
	this.m_bFilters = false;
	this.m_sFiltername = "blendtrans";
	
	this.setFiltername = function( sFiltername )
	{
		this.m_sFiltername = sFiltername;
	};
}

CSlideScreenTransition.prototype = new CSlideScreen;
CSlideScreenTransition.prototype.constructor = CSlideScreenTransition;


CSlideScreenTransition.prototype.initialize
= function()
{
	if ( CSlideScreen.prototype.initialize.call( this ) )
	{
		if ( document.all )
		{
			if ( typeof(this.m_oDiv.filters) != "undefined" )
			{
				if ( 0 < this.m_oDiv.filters.length )
					this.m_bFilters = true;
			}
		}
		return true;
	}
	else
	{
		return false;
	}
};


CSlideScreenTransition.prototype.hasFilters
= function()
{
	return this.m_bFilters;
};


CSlideScreenTransition.prototype.getTransitionDelay
= function()
{
	if ( this.m_bFilters )
	{
		if ( typeof( this.m_oDiv.filters[this.m_sFiltername].duration ) == "undefined" )
			return 0;
		else
			return this.m_oDiv.filters[this.m_sFiltername].duration;		// return as seconds
	}
	else
	{
		return 0;
	}
};


CSlideScreenTransition.prototype.getTransitionStatus
= function()
{
	if ( this.m_bFilters )
	{
		if ( "undefined" == typeof(this.m_oDiv.filters[this.m_sFiltername].status) )
			return 0;
		else
			return this.m_oDiv.filters[this.m_sFiltername].status;
	}
	else
	{
		return 0;
	}
};


CSlideScreenTransition.prototype.displaySlidePreprocess
= function( sSlideValue )
{
	//window.alert( this.m_sDivID );
	CSlideScreen.prototype.displaySlidePreprocess( this, sSlideValue );
	if ( 0 < this.getTransitionDelay() )
	{
		//this.m_oDiv.filters[0].transition = getRandomTransition();
		this.m_oDiv.filters[this.m_sFiltername].apply();
	}
};

CSlideScreenTransition.prototype.displaySlidePostprocess
= function( sSlideValue )
{
	if ( 0 < this.getTransitionDelay() )
		this.m_oDiv.filters[this.m_sFiltername].play();
	CSlideScreen.prototype.displaySlidePostprocess( this, sSlideValue );
};



//	class CSlideScreenTransitionRandom : CSlideScreenTransition

function CSlideScreenTransitionRandom( sDivID )
{
	CSlideScreenTransition.call( this, sDivID );
	this.setFiltername( "revealtrans" );
}

CSlideScreenTransitionRandom.prototype = new CSlideScreenTransition;
CSlideScreenTransitionRandom.prototype.constructor = CSlideScreenTransitionRandom;


CSlideScreenTransitionRandom.prototype.displaySlidePreprocess
= function( sSlideValue )
{
	if ( 0 < this.getTransitionDelay() )
	{
		this.m_oDiv.filters[this.m_sFiltername].transition = this.getRandomTransition();
	}
	CSlideScreenTransition.prototype.displaySlidePreprocess.call( this, sSlideValue );
};



CSlideScreenTransitionRandom.prototype.getRandomTransitionAvoid
= function( n )
{
	var bFound = false;
	
	switch ( n )
	{
	case 12:	// Random dissolve
	case 21:	// Random bars horizontal
	case 22:	// Random bars vertical
	case 23:	// Random transition from above possible values
		bFound = true;
		break;
	default:
		break;
	}
	return bFound;
};

CSlideScreenTransitionRandom.prototype.getRandomTransition
= function()
{
	var i;
	do
	{
		i = Math.round(Math.random() * 23);
	} while ( this.getRandomTransitionAvoid( i ) );	// we want to avoid random dissolve because it is slow
	return i;
};









// class CSlideProjector

function CSlideProjector()
{
	this.m_oScreen = null;
	this.m_oCarousel = null;

	
	
	this.setScreen = function( oScreen )
	{
		this.m_oScreen = oScreen;
	};


	this.setCarousel = function( oCarousel )
	{
		this.m_oCarousel = oCarousel;
	};
	
	
	this.setTimeout = function( sMethod, nSeconds )
	{
		var tempThis = this;
		var funcRef = (function(){tempThis[sMethod]()});
		return setTimeout( funcRef, nSeconds * 1000 );
	};
	
}


CSlideProjector.prototype.run
= function()
{
};





//	class CSlideProjectorShowSingle : CSlideProjector

function CSlideProjectorShowSingle()
{
	CSlideProjector.call( this );
	
	this.m_nCurrentSlide = 0;
	this.m_nBlendDelay = 0;
	this.m_dPauseTarget = new Date();
	
	this.m_nPicturePause = 7.0;		// number of seconds to display picture before transition
	
	this.setPicturePause = function( nPicturePause )
	{
		this.m_nPicturePause = nPicturePause;
	};
}

CSlideProjectorShowSingle.prototype = new CSlideProjector;
CSlideProjectorShowSingle.prototype.constructor = CSlideProjectorShowSingle;



CSlideProjectorShowSingle.prototype.run
= function()
{

	this.m_nBlendDelay = this.m_oScreen.getTransitionDelay();
	
	var oCanvas = this.m_oScreen.getCanvas();
	var tempThis = this;
	oCanvas.onfilterchange = (function(e){e = e || window.event;tempThis["handleTransition"](e,this)});
	//oCanvas.onfilterchange = CSlideProjector.g_aInstances[this.m_idInstance].handleTransition;
	
	this.m_nCurrentSlide = 0;
	//this.slideShowNextImage();
	this.setTimeout( "makeReadyForShow", 0.01 );

};



CSlideProjectorShowSingle.prototype.makeReadyForShow
= function()
{
	var oSlide = this.m_oCarousel.item( this.m_nCurrentSlide );
	if ( oSlide )
	{
		if ( oSlide.precook() )
			this.setTimeout( "slideShowNextImage", 0.01 );
		else
			this.setTimeout( "makeReadyForShow", 0.125 );
	}
	else
	{
		this.setTimeout( "slideShowNextImage", 0.01 );
	}
};




CSlideProjectorShowSingle.prototype.slideShowNextImage
= function()
{
	//window.status = "slideShowNextImage";
	var o = this.nextSlide();
	var contentcontainer;
	contentcontainer = o.projectSlide( this.m_oScreen.height(), this.m_oScreen.width() );
	
	this.m_oScreen.displaySlide( contentcontainer );
	
	if ( 0 == this.m_nBlendDelay )
		this.setTimeout( "waitOnTransition", 0.1 );
	
	//window.alert( "delay = " + this.m_nBlendDelay + "\n" + "pause = " + this.m_nPicturePause  );
	
};


CSlideProjectorShowSingle.prototype.nextSlide
= function()
{
	var n = this.m_nCurrentSlide;
	var o = this.m_oCarousel.item( n );
	this.m_nCurrentSlide = (n + 1) % this.m_oCarousel.length();
	return o;
};


CSlideProjectorShowSingle.prototype.handleTransition
= function( e )
{
	//window.status = "handleTransition";
	if ( 0 == this.m_oScreen.getTransitionStatus() )
	{
		this.setTimeout("pauseBegin", 0.1 );
		this.outsideEvent( e );
	}
	return true;
};

CSlideProjectorShowSingle.prototype.outsideEvent
= function( e )
{
	// function for override
};


CSlideProjectorShowSingle.prototype.waitOnTransition
= function()
{
	//window.status = "waitOnTransition";
	if ( 0 < this.m_nBlendDelay )
	{
		if ( 0 == this.m_oScreen.getTransitionStatus() )
			this.setTimeout("pauseBegin", 0.1 );
		else
			this.setTimeout("waitOnTransition", 0.25 );
	}
	else
	{
		this.setTimeout("pauseBegin", 0.1 );
	}
};



CSlideProjectorShowSingle.prototype.pauseBegin
= function()
{
	//window.status = "pauseBegin";
	var d = new Date();
	this.m_dPauseTarget = dateAddSeconds( d, this.m_nPicturePause );
	this.setTimeout( "pauseShow", this.m_nPicturePause );
};

CSlideProjectorShowSingle.prototype.pauseShow
= function()
{
	//window.status = "pauseShow";
	var d = new Date();
	if ( d < this.m_dPauseTarget )
	{
		var t = d.getTime();
		var tt = this.m_dPauseTarget.getTime();
		this.setTimeout( "pauseShow", (tt-t)/1000 );
	}
	else
	{
		this.setTimeout( "makeReadyForShow", 0.01 );
	}
};




//	class CSlideCardFactory
//		encapsulates the intelligence of which type of CSlideCard to create
//		The factory will actually create the slide-card, initialize it, and then
//		return it for use in the carousel.
//		NOTE: Initializing is NOT the same as precooking.

function CSlideCardFactory()
{
}

CSlideCardFactory.prototype.makeSlide
= function( s )
{
	var sPrefix = s.substr(0,4);
	if ( "htm:" == sPrefix )
		return new CSlideHTML( s.substr(4) );
	else if ( "img:" == sPrefix )
		return new CSlidePicture( s.substr(4) );
	else
		return new CSlidePicture( s );
}





// class CSlideCarousel
//		encapsulates the list of slides to be presented
//		operates as a factory for CSlidePicture based on the received array

function CSlideCarousel()
{
	this.m_aSlideCards = new Array();
	this.m_oSlideFactory = new CSlideCardFactory();
	this.m_nCooker = -1;
	this.m_nCookerTimer = 0.5;	// seconds delay between precook
	
	this.setTimeout = function( sMethod, nSeconds )
	{
		var tempThis = this;
		var funcRef = (function(){tempThis[sMethod]()});
		return setTimeout( funcRef, nSeconds * 1000 );
	};
	
	this.setSlideFactory = function( oFactory )
	{
		this.m_oSlideFactory = oFactory;
	};
	
	this.setCookerTimer = function( n )
	{
		this.m_nCookerTimer = n;
	};
	
	this.length = function()
	{
		return this.m_aSlideCards.length;
	};
	
	this.item = function( nIndex )
	{
		var n = nIndex % this.m_aSlideCards.length;
		return this.m_aSlideCards[n];
	};
	
	this.appendItem = function( oCard )
	{
		var n = this.m_aSlideCards.length;
		this.m_aSlideCards[n] = oCard;
		return n;
	};
	
}





CSlideCarousel.prototype.load 
= function( aList )
{
	var	nLen = aList.length;
	if ( 0 < nLen )
	{
		m_aSlideCards = new Array( nLen-1 );
		if ( this.m_aSlideCards )
		{
			var sPrefix;
			var	i;
			for ( i = 0; i < aList.length; ++i )
			{
				this.m_aSlideCards[i] = this.m_oSlideFactory.makeSlide( aList[i] );
			}
		}
	}
};




CSlideCarousel.prototype.precookCards
= function()
{
	this.m_nCooker ++;
	if ( this.m_nCooker < this.m_aSlideCards.length )
	{
		this.m_aSlideCards[this.m_nCooker].precook();
		this.setTimeout( "precookCards", this.m_nCookerTimer );
	}
};








// class CSlideCard
//		abstract class for items that will be represented on the screen
//		holds the actual source string

function CSlideCard( sSource )
{
	this.m_source = sSource;
}

CSlideCard.prototype.type
= function()
{
	return "unknown";
};


CSlideCard.prototype.precook
= function()
{
	// the default base function does nothing
	return true;
};

CSlideCard.prototype.projectSlide
= function( nHeight, nWidth )
{
	return "";	// return nothing for the default function
};







// class CSlidePicture : CSlideCard
//		encapsulates the image or representation that will be presented on the screen
//		holds the actual string to the picture as well as the cached object


function CSlidePicture( sSource )
{
	CSlideCard.call( this, sSource );
	
	this.m_bCached = false;
	this.m_oImage = new Image();
}

CSlidePicture.prototype = new CSlideCard;
CSlidePicture.prototype.constructor = CSlidePicture;


CSlidePicture.prototype.type
= function()
{
	return "picture";
};

CSlidePicture.prototype.height
= function()
{
	if ( this.precook() )
		return this.m_oImage.height;
	else
		return 0;
};

CSlidePicture.prototype.width
= function()
{
	if ( this.precook() )
		return this.m_oImage.width;
	else
		return 0;
};


CSlidePicture.prototype.precook
= function()
{
	var	bResult = true;
	if ( ! this.m_bCached )
	{
		this.m_oImage.src = this.m_source;
		this.m_bCached = true;
	}
	if ( ! this.m_oImage.complete )
		bResult = false;
	return bResult;
};




CSlidePicture.prototype.projectSlide
= function( nHeight, nWidth )
{
	var h = this.height();
	var w = this.width();
	
	if ( 0 < h  &&  0 < w )
	{
		var WidthRatio  = w / nWidth;
		var HeightRatio = h / nHeight;
		var Ratio = WidthRatio > HeightRatio ? WidthRatio : HeightRatio;
		
	
		w  = Math.floor(w / Ratio);
		h = Math.floor(h / Ratio);
		
		if ( nHeight < h )
			h = nHeight;
		if ( nWidth < w )
			w = nWidth;
	}
	else
	{
		h = nHeight;
		w = nWidth;
	}
	
	var s = '<div class="SlideShowImg"><img src="'+this.m_source+'" border="0" height="'+h+'" width="'+w+'" ></div>';

	return s;
};






// class CSlideHTML : CSlideCard

function CSlideHTML( sSource )
{
	CSlideCard.call( this, sSource );
}

CSlideHTML.prototype = new CSlideCard;
CSlideHTML.prototype.constructor = CSlideHTML;


CSlideHTML.prototype.type
= function()
{
	return "html";
};



CSlideHTML.prototype.projectSlide
= function( nHeight, nWidth )
{
	// for HTML all we need to do is return the source
	return '<div class="SlideShowHtm">' + this.m_source + '</div>';
};







///////////////////////////////////////////////////////////
//  Library routines
///////////////////////////////////////////////////////////


function dateAddSeconds( d, s )
{
	var dd = new Date( d );
	var m = 0;
	var n;
	var i;
	var	r;
	i = Math.floor( s );
	r = (s - i) * 1000;	// get milliseconds
	if ( 0 < r )
	{
		n = dd.getMilliseconds();
		n += r;
		if ( 999 < n )
		{
			++i;
			n = n % 1000;
		}
		r = n;
	}
	else
	{
		r = dd.getMilliseconds();
	}
	m = 0;
	if ( 0 < i )
	{
		n = dd.getSeconds();
		n += i;
		if ( 59 < n )
		{
			m += Math.floor(n / 60);
			n = n % 60;
		}
		dd.setSeconds( n );
	}
	if ( 0 < r )
		dd.setMilliseconds( r );

	i = m;
	m = 0;
	if ( 0 < i )
	{
		n = dd.getMinutes();
		n += i;
		if ( 59 < n )
		{
			m += Math.floor(n / 60);
			n = n % 60;
		}
		dd.setMinutes( n );
	}
	i = m;
	m = 0;
	if ( 0 < i )
	{
		n = dd.getHours();
		n += i;
		if ( 23 < n )
		{
			m += Math.floor(n / 24);
			n = n % 24;
		}
		dd.setHours( n );
	}
	
	// handle one strange case
	if ( dd < d )
	{
		dd.setTime( d.getTime() + (s*1000) );
	}
	
	return dd;
	
}



//	Return a reference to an anonymous inner function created
//	with a function expression:-

function slideFuncRef( obj, sMethodName, sArgs)
{
    /* This inner function is to be executed with - setTimeout
       - and when it is executed it can read, and act upon, the
       parameters passed to the outer function:-
    */
    return (function(){
        obj[sMethodName]();
    });
}


function getDivHeight( o )
{
	var h;
	
	if ( o.offsetHeight )
		h = o.offsetHeight;
	else if ( o.style.pixelHeight )
		h = o.style.pixelHeight;
	else if ( o.height )
		h = o.height;

	return h;
}


function getDivWidth( o )
{
	var w;
	if ( o.offsetWidth )
		w = o.offsetWidth;
	else if ( o.style.pixelWidth )
		w = o.style.pixelWidth;
	else if ( o.width )
		w = o.width;
	return w;
}





