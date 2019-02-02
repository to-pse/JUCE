/*
  ==============================================================================

   This file is part of the JUCE library.
   Copyright (c) 2017 - ROLI Ltd.

   JUCE is an open source library subject to commercial or open-source
   licensing.

   The code included in this file is provided under the terms of the ISC license
   http://www.isc.org/downloads/software-support-policy/isc-license. Permission
   To use, copy, modify, and/or distribute this software for any purpose with or
   without fee is hereby granted provided that the above copyright notice and
   this permission notice appear in all copies.

   JUCE IS PROVIDED "AS IS" WITHOUT ANY WARRANTY, AND ALL WARRANTIES, WHETHER
   EXPRESSED OR IMPLIED, INCLUDING MERCHANTABILITY AND FITNESS FOR PURPOSE, ARE
   DISCLAIMED.

   Extension of the JUCE library to support createNewDevice on Windows platform
   Created by Tobias Erichsen, usage restrictions apply

   Required teVirtualMIDI version:  1.2.11

   This file needs to be included into juce_win32_Midi.cpp

   Necessary defines in Projucer Preprocessor Definitions:
   COMMERCIAL_LICENSE_TE_VIRTUAL_MIDI
   (TE_VIRTUAL_MIDI_MAX_SYSEX_SIZE=<max sysex size>)
  ==============================================================================
*/


#if defined(COMMERCIAL_LICENSE_TE_VIRTUAL_MIDI)
/* the file below is part of the teVirtualMIDI sdk available here: http://www.tobias-erichsen.de/software/virtualmidi/virtualmidi-sdk.html */
#include "teVirtualMIDI.h"

#ifndef TE_VIRTUAL_MIDI_MAX_SYSEX_SIZE
#define TE_VIRTUAL_MIDI_MAX_SYSEX_SIZE 8192
#endif

/* the teVirtualMIDI.dll C interface used by this wrapper */
extern "C"
{
	typedef LPVM_MIDI_PORT	( CALLBACK *VIRTUALMIDICREATEPORTEX3 )	( LPCWSTR portName, LPVM_MIDI_DATA_CB callback, DWORD_PTR dwCallbackInstance, DWORD maxSysexLength, DWORD flags, GUID *manufacturer, GUID *product );
	typedef void			( CALLBACK *VIRTUALMIDICLOSEPORT )		( LPVM_MIDI_PORT midiPort );
	typedef BOOL			( CALLBACK *VIRTUALMIDISENDDATA )		( LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length );
}


/* singleton class to dynamically load the teVirtualMIDI.dll */
class VirtualMidiDll
{
public:
	static VirtualMidiDll& getInstance()
	{
		static VirtualMidiDll    instance; 
		return instance;
	}
	VirtualMidiDll( VirtualMidiDll const& ) = delete;
	void operator=( VirtualMidiDll const& ) = delete;

	DWORD			getError	();

	/* the wrapped calls to the teVirtualMIDI.dll */
	LPVM_MIDI_PORT	create		( LPCWSTR portName, LPVM_MIDI_DATA_CB callback, DWORD_PTR dwCallbackInstance, DWORD maxSysexLength, DWORD flags, GUID *manufacturer, GUID *product );
	void			close		( LPVM_MIDI_PORT midiPort );
	bool			send		( LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length );

private:
	HMODULE						hModule						= nullptr;
	DWORD						error						= 0;
	VIRTUALMIDICREATEPORTEX3	virtualMIDICreatePortEx3	= nullptr;
	VIRTUALMIDICLOSEPORT		virtualMIDIClosePort		= nullptr;
	VIRTUALMIDISENDDATA			virtualMIDISendData			= nullptr;

	VirtualMidiDll();
	~VirtualMidiDll();
};


/* dll singleton implementation */
VirtualMidiDll::VirtualMidiDll()
{
	SetErrorMode(SEM_NOOPENFILEERRORBOX | SEM_FAILCRITICALERRORS);

	this->virtualMIDICreatePortEx3 = nullptr;
	this->virtualMIDIClosePort = nullptr;
	this->virtualMIDISendData = nullptr;

	this->hModule = LoadLibraryA( "teVirtualMIDI.dll" );
	if( !this->hModule )
	{
		this->error = GetLastError();
		return;
	}
	this->virtualMIDICreatePortEx3 = ( VIRTUALMIDICREATEPORTEX3 )GetProcAddress( this->hModule, "virtualMIDICreatePortEx3" );
	if( !this->virtualMIDICreatePortEx3 )
	{
		this->error = GetLastError();
	}
	this->virtualMIDIClosePort = ( VIRTUALMIDICLOSEPORT )GetProcAddress( this->hModule, "virtualMIDIClosePort" );
	if( !this->virtualMIDIClosePort )
	{
		this->error = GetLastError();
	}
	this->virtualMIDISendData = ( VIRTUALMIDISENDDATA )GetProcAddress( this->hModule, "virtualMIDISendData" );
	if( !this->virtualMIDISendData )
	{
		this->error = GetLastError();
	}

	if( !this->virtualMIDICreatePortEx3 || !this->virtualMIDIClosePort || !this->virtualMIDISendData )
	{
		FreeLibrary( this->hModule );
		this->hModule = nullptr;
		this->virtualMIDICreatePortEx3 = nullptr;
		this->virtualMIDIClosePort = nullptr;
		this->virtualMIDISendData = nullptr;
	}
}

VirtualMidiDll::~VirtualMidiDll()
{
	if( this->hModule )
	{
		FreeLibrary( this->hModule );
		this->hModule = nullptr;
		this->virtualMIDICreatePortEx3 = nullptr;
		this->virtualMIDIClosePort = nullptr;
		this->virtualMIDISendData = nullptr;
	}
}

DWORD VirtualMidiDll::getError()
{
	return this->error;
}

LPVM_MIDI_PORT VirtualMidiDll::create( LPCWSTR portName, LPVM_MIDI_DATA_CB callback, DWORD_PTR dwCallbackInstance, DWORD maxSysexLength, DWORD flags, GUID *manufacturer, GUID *product )
{

	if( !this->virtualMIDICreatePortEx3 )
	{
		return nullptr;
	}
	return ( *this->virtualMIDICreatePortEx3 )( portName, callback, dwCallbackInstance, maxSysexLength, flags, manufacturer, product );
}

void VirtualMidiDll::close( LPVM_MIDI_PORT port )
{

	if( !this->virtualMIDIClosePort )
	{
		return;
	}
	this->virtualMIDIClosePort( port );
}

bool VirtualMidiDll::send( LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length )
{

	if( this->virtualMIDISendData )
	{
		return this->virtualMIDISendData( midiPort, midiDataBytes, length ) != FALSE;
	}

	return false;
}


namespace juce {

	/* the actual callback for receiving MIDI-data from a created port */
	extern "C" void CALLBACK midiDataCallback( LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length, DWORD_PTR dwCallbackInstance );


	class Win32VirtualMidi;

	/* The MidiInput subclass to be used for MidiInput::createNewDevice*/
	class Win32VirtualMidiInput : public MidiInput
	{
	public:
		Win32VirtualMidiInput(Win32VirtualMidi *port, MidiInputCallback* callback);

		~Win32VirtualMidiInput();
		void startInt();
		void stopInt();
	private:
		friend void CALLBACK midiDataCallback(LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length, DWORD_PTR dwCallbackInstance);
		Win32VirtualMidi			*port;
		juce::MidiInputCallback		*callback = nullptr;
		bool						 started = false;
		explicit Win32VirtualMidiInput(const String& portName) : Win32VirtualMidiInput(nullptr, nullptr) { (void)portName; };
	};

	/* The MidiOutput subclass to be used for MidiOutput::createNewDevice*/
	class Win32VirtualMidiOutput : public MidiOutput
	{
	public:
		Win32VirtualMidiOutput(Win32VirtualMidi *port);
		~Win32VirtualMidiOutput();
		void sendMessageNowInt(const MidiMessage& message);
	private:
		Win32VirtualMidi			*port;
		explicit Win32VirtualMidiOutput(const String& portName) : Win32VirtualMidiOutput(nullptr) { (void)portName; };
	};


	class Win32VirtualMidi
	{
	public:
		Win32VirtualMidi(const String& portName, DWORD maxSysexLength, const Uuid *manufacturer, const Uuid *product);
		~Win32VirtualMidi();

		static Win32VirtualMidiInput *bindMidiInput(const String &portName, MidiInputCallback* callback, DWORD maxSysexLength = TE_VIRTUAL_MIDI_MAX_SYSEX_SIZE, const Uuid *manufacturer = nullptr, const Uuid *product = nullptr );
		static Win32VirtualMidiOutput *bindMidiOutput(const String &portName, DWORD maxSysexLength = TE_VIRTUAL_MIDI_MAX_SYSEX_SIZE, const Uuid *manufacturer = nullptr, const Uuid *product = nullptr);
	private:
		VirtualMidiDll&				 dll = VirtualMidiDll::getInstance();
		DWORD						 error = 0;
		LPVM_MIDI_PORT				 port = nullptr;
		String						 name;
		DWORD						 bindsIn = 0;
		DWORD						 bindsOut = 0;

		/* per port list of callbacks */
		Array<Win32VirtualMidiInput *> midiIns;
		Array<Win32VirtualMidiOutput *> midiOuts;

		friend Win32VirtualMidiInput;
		friend Win32VirtualMidiOutput;
		friend void CALLBACK midiDataCallback(LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length, DWORD_PTR dwCallbackInstance);

		/* static map of all open ports referencable via name */
		static HashMap<String, Win32VirtualMidi*>& getPorts()
		{
			// store a weak reference to the shared key windows
			static HashMap<String, Win32VirtualMidi*> ports;
			return ports;
		}
		static void removePort(const String &portName);
		static void addPort(const String &portName, Win32VirtualMidi *port);
		static Win32VirtualMidi *getPort(const String &portName);

	};

	Win32VirtualMidi::Win32VirtualMidi(const String& portName, DWORD maxSysexLength, const Uuid *manufacturer, const Uuid *product) : name(portName)
	{
		GUID	*manu = nullptr;
		GUID	*prod = nullptr;

		if (manufacturer)
		{
			manu = (GUID *)manufacturer->getRawData();
		}

		if (product)
		{
			prod = (GUID *)product->getRawData();
		}

		this->port = this->dll.create(portName.toWideCharPointer(), midiDataCallback, (DWORD_PTR)this, maxSysexLength, 0, manu, prod);
		if (!this->port)
		{
			this->error = GetLastError();
		}
	}

	Win32VirtualMidi::~Win32VirtualMidi()
	{
		if (this->port)
		{
			this->dll.close(this->port);
		}
	}

	void Win32VirtualMidi::removePort(const String &portName)
	{
		Win32VirtualMidi::getPorts().remove(portName);
	}

	void Win32VirtualMidi::addPort(const String &portName, Win32VirtualMidi *port)
	{
		Win32VirtualMidi::getPorts().set(portName, port);
	}

	Win32VirtualMidi *Win32VirtualMidi::getPort(const String &portName)
	{
		return Win32VirtualMidi::getPorts()[portName];
	}



	Win32VirtualMidiInput *Win32VirtualMidi::bindMidiInput(const String &portName, MidiInputCallback* callback, DWORD maxSysexLength, const Uuid *manufacturer, const Uuid *product)
	{
		Win32VirtualMidi		*port = nullptr;
		Win32VirtualMidiInput	*result = nullptr;

		port = Win32VirtualMidi::getPort(portName);
		if (!port)
		{
			port = new Win32VirtualMidi(portName, maxSysexLength, manufacturer, product);
			if (port->error)
			{
				delete port;
				port = nullptr;
			}
			else
			{
				addPort(portName, port);
			}
		}
		if (port)
		{
			result = new Win32VirtualMidiInput(port, callback);
		}
		return result;
	}

	Win32VirtualMidiOutput *Win32VirtualMidi::bindMidiOutput(const String &portName, DWORD maxSysexLength, const Uuid *manufacturer, const Uuid *product)
	{
		Win32VirtualMidi		*port = nullptr;
		Win32VirtualMidiOutput	*result = nullptr;

		port = Win32VirtualMidi::getPort(portName);
		if (!port)
		{
			port = new Win32VirtualMidi(portName, maxSysexLength, manufacturer, product);
			if (port->error)
			{
				delete port;
				port = nullptr;
			}
			else
			{
				addPort(portName, port);
			}
		}
		if (port)
		{
			result = new Win32VirtualMidiOutput(port);
		}
		return result;
	}

	/* the Win32VirtualMidiInput implementation */
	Win32VirtualMidiInput::Win32VirtualMidiInput(Win32VirtualMidi *port, MidiInputCallback* callback): MidiInput( port->name )
	{

		this->port = port;
		this->callback = callback;
		port->midiIns.add(this);
	}

	Win32VirtualMidiInput::~Win32VirtualMidiInput()
	{
		if( this->port )
		{
			this->port->midiIns.removeFirstMatchingValue(this);
			if (this->port->midiIns.isEmpty() && this->port->midiOuts.isEmpty())
			{
				Win32VirtualMidi::removePort(this->port->name);
			}
		}
	}

	void Win32VirtualMidiInput::startInt()
	{
		this->started = true;
	}

	void Win32VirtualMidiInput::stopInt()
	{
		this->started = false;
	}

	/* the teVirtualMIDI callback for incoming MIDI-data */
	void CALLBACK midiDataCallback( LPVM_MIDI_PORT midiPort, LPBYTE midiDataBytes, DWORD length, DWORD_PTR dwCallbackInstance )
	{
		(void)midiPort;            // Unused
		MidiMessage	msg = MidiMessage(midiDataBytes, length);

		auto* port = reinterpret_cast<Win32VirtualMidi*> ( dwCallbackInstance );

		for (int i = 0; i < port->midiIns.size(); i++)
		{
			if (port->midiIns[i]->callback)
			{
				port->midiIns[i]->callback->handleIncomingMidiMessage(port->midiIns[i], msg);
			}
		}

	}


	/* the Win32VirtualMidiOutput implementation */
	Win32VirtualMidiOutput::Win32VirtualMidiOutput(Win32VirtualMidi *port) : MidiOutput( port->name )
	{
		this->port = port;
		port->midiOuts.add(this);
	}

	Win32VirtualMidiOutput::~Win32VirtualMidiOutput()
	{
		if (this->port)
		{
			this->port->midiOuts.removeFirstMatchingValue(this);
			if (this->port->midiIns.isEmpty() && this->port->midiOuts.isEmpty())
			{
				Win32VirtualMidi::removePort(this->port->name);
			}
		}
	}

	void Win32VirtualMidiOutput::sendMessageNowInt( const MidiMessage& message )
	{
		if( !this->port )
		{
			return;
		}
		this->port->dll.send( this->port->port, (LPBYTE)message.getRawData(), message.getRawDataSize() );
	}


}

#endif /* !COMMERCIAL_LICENSE_TE_VIRTUAL_MIDI */