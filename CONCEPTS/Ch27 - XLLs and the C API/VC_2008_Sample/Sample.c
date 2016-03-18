#include <windows.h>
#include "xlcall.h"

#ifdef __cplusplus
#define EXPORT extern "C" __declspec(dllexport)
#else
#define EXPORT __declspec(dllexport)
#endif

#define NUM_FUNCTIONS		3		// number of functions in the table
#define NUM_REGISTER_ARGS	11		// number of register arguments for each function
#define MAX_LENGTH			255		// maximum allowable string length for a register argument

static char gszFunctionTable[NUM_FUNCTIONS][NUM_REGISTER_ARGS][MAX_LENGTH] =
{
	{" AddTwo",                     // procedure
     " BBB",                        // type_text
     " AddTwo",                     // function_text
     " d1, d2",                     // argument_text
     " 1",                          // macro_type
     " Sample Add-In",              // category
     " ",                           // shortcut_text
     " ",                           // help_topic
     " Adds the two arguments.",	// function_help
     " The first number to add.",   // argument_help1
     " The second number to add."	// argument_help2
    },
    {" MultiplyTwo",
     " BBB",
     " MultiplyTwo",
     " d1, d2",
     " 1",
     " Sample Add-In",
     " ",
     " ",
     " Multiplies the two arguments.",
     " The first number to multiply.",
     " The second number to multiply."
    },
    {" IFERROR",
     " RRR",
     " IFERROR",
     " ToEvaluate, Default",
     " 1",
     " Sample Add-In",
     " ",
     " ",
     " If the first argument is an error value, the second "
		"argument is returned. Otherwise the first argument "
		"is returned.",
     " The argument to be checked for an error condition.",
     " The value to return if the first argument is an error."
    }
};

void HandleRegistration(BOOL);


// Windows expects this function to be present in any Windows DLL,
// therefore we will provide it, but not make use of it.
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD fdwReason, PVOID pvReserved)
{
	return TRUE;
}


EXPORT int WINAPI xlAutoOpen(void)
{
	static XLOPER xDLL;
	int i, j;   

	// In the following loop, the strings in 
	// gszFunctionTable are byte-counted.
	for (i = 0; i < NUM_FUNCTIONS; ++i)
		for (j = 0; j < NUM_REGISTER_ARGS; ++j)
			gszFunctionTable[i][j][0] = 
				(BYTE) lstrlenA(gszFunctionTable[i][j] + 1);

	// Register the functions using our custom procedure.
	HandleRegistration(TRUE);

	return 1;
}


EXPORT int WINAPI xlAutoClose(void)
{
	// Unregister the worksheet functions 
	// using our custom procedure.
	HandleRegistration(FALSE);
	return 1;
}


EXPORT LPXLOPER WINAPI xlAddInManagerInfo(LPXLOPER xlAction)
{
	static XLOPER xlReturn, xlLongName, xlTemp;

	// Coerce the argument XLOPER to an integer.
	xlTemp.xltype = xltypeInt;
	xlTemp.val.w = xltypeInt;
	Excel4(xlCoerce, &xlReturn, 2, xlAction, &xlTemp);

	// The only valid argument value is 1. In this case we 
	// return the long name for the XLL. Any other value should 
	// result in the return of a #VALUE! error.
	if(1 == xlReturn.val.w)
	{
		xlLongName.xltype = xltypeStr;
		xlLongName.val.str = "\021Sample XLL Add-In";
	}
	else
	{
		xlLongName.xltype = xltypeErr;
		xlLongName.val.err = xlerrValue;
	}

	return &xlLongName;
}


////////////////////////////////////////////////////////////////
// Comments:	This function handles registering and 
//				unregistering all of the custom worksheet 
//				functions specified in our function table.
//				
// Parameters:	bRegister	[in] Pass TRUE to register all the 
//							custom worksheet functions or FALSE
//							to unregister them.
//
// Returns:		No return.
//
static void HandleRegistration(BOOL bRegister)
{
	XLOPER	xlXLLName, xlRegID, xlRegArgs[NUM_REGISTER_ARGS];
	int		i, j;

	// Get the filename of the XLL by calling xlGetName.
	Excel4(xlGetName, &xlXLLName, 0);

	// All of the XLOPER arguments passed to the Register
	// function will have the type xltypeStr.
	for (i = 0; i < NUM_REGISTER_ARGS; ++i)
		xlRegArgs[i].xltype = xltypeStr;

	for (i = 0; i < NUM_FUNCTIONS; ++i)
	{
		// Load the XLOPER arguments to the Register function.
		for(j = 0; j < NUM_REGISTER_ARGS; ++j)
			xlRegArgs[j].val.str = gszFunctionTable[i][j];

		if (TRUE == bRegister)
		{
			// Register each function.
			// NOTE: The number of xlRegArgs[] arguments passed
			// here must be equal to NUM_REGISTER_ARGS - 1.
			Excel4(xlfRegister, 0, NUM_REGISTER_ARGS + 1, 
				&xlXLLName,
				&xlRegArgs[0], &xlRegArgs[1], &xlRegArgs[2], 
				&xlRegArgs[3], &xlRegArgs[4], &xlRegArgs[5], 
				&xlRegArgs[6], &xlRegArgs[7], &xlRegArgs[8], 
				&xlRegArgs[9], &xlRegArgs[10]);
		}
		else
		{
			// Unregister each function.
			// Due to a bug in Excel's C API this is a two-step
			// process. Thanks to Laurent Longre for discovering
			// the workaround described here.
			// Step 1: Redefine each custom worksheet function 
			// as a hidden function (change the macro_type 
			// argument to 0).
			xlRegArgs[4].val.str = "\0010";
			// Step 2: Re-register each function as a hidden 
			// function.
			// NOTE: The number of xlRegArgs[] arguments passed
			// here must be equal to NUM_REGISTER_ARGS - 1.
			Excel4(xlfRegister, 0, NUM_REGISTER_ARGS + 1, 
				&xlXLLName,
				&xlRegArgs[0], &xlRegArgs[1], &xlRegArgs[2],
				&xlRegArgs[3], &xlRegArgs[4], &xlRegArgs[5],
				&xlRegArgs[6], &xlRegArgs[7], &xlRegArgs[8],
				&xlRegArgs[9], &xlRegArgs[10]);
			// Step 3: Unregister the now hidden function.
			// First, get the Register ID for the function. 
			// Since xlfRegisterId will return a non-pointer
			// type to the xlRegID XLOPER, we do not need to
			// call xlFree on it.
			Excel4(xlfRegisterId, &xlRegID, 2, &xlXLLName, 
												&xlRegArgs[0]);
			// Second, unregister the function using its
			// Register ID.
			Excel4(xlfUnregister, 0, 1, &xlRegID);
		}
	}

	// Since xlXLLName holds a pointer that is managed by Excel,
	// we must call xlFree on it.
	Excel4(xlFree, 0, 1, &xlXLLName);
}


EXPORT double WINAPI AddTwo(double d1, double d2)
{
	return d1 + d2;
}


EXPORT double WINAPI MultiplyTwo(double d1, double d2)
{
	return d1 * d2;
}


////////////////////////////////////////////////////////////////
// Comments:	This function provides a short-cut replacement
//				for the common worksheet function construct:
//				=IF(ISERROR(<some_function>),0,<some_function>)
//
// Arguments:	ToEvaluate	[in] A value, expression or cell
//							reference to be evaluated.
//				Default		[in] A value, expression or cell 
//							reference to be returned if the 
//							ToEvaluate argument evaluates to an
//							error condition.
//
// Returns:		ToEvaluate if not an error, Default otherwise.
//
EXPORT LPXLOPER IFERROR(LPXLOPER ToEvaluate, LPXLOPER Default)
{
	int				IsError = 0;
	XLOPER			xlResult;
	static XLOPER	xlBadArgErr;

	// This is the return value for bad or missing arguments.
	xlBadArgErr.xltype = xltypeErr;
	xlBadArgErr.val.err = xlerrValue;

	// Check for missing arguments.
	if ((xltypeMissing == ToEvaluate->xltype) || 
		(xltypeMissing == Default->xltype))
		return &xlBadArgErr;
	
	switch (ToEvaluate->xltype)
	{
		// The first four all indicate valid ToEvaluate types.
		// Drop out and use ToEvaluate as the return value.
		case xltypeNum:
		case xltypeStr:
		case xltypeBool:
		case xltypeInt:
			break;
		// A cell reference must be dereferenced to see what it
		// contains.
		case xltypeSRef:
		case xltypeRef:
			if (xlretUncalced == Excel4(xlCoerce, &xlResult, 1,
													ToEvaluate))
				// If we're looking at an uncalculateded cell,
				// return immediately. Excel will call this
				// function again once the dependency has been 
				// calculated.
				return 0;
			else
			{
				if (xltypeMulti == xlResult.xltype)
					// Multi-cell arguments are not permitted.
					return &xlBadArgErr;	
				else if (xltypeErr == xlResult.xltype)
					// ToEvaluate is a single cell containing an
					// error. Return Default instead.
					IsError = 1;			
			}
			// ToEvaluate is returned for all other types.
			// Always call xlFree on the return value from
			// Excel4.
			Excel4(xlFree, 0, 1, &xlResult);
			break;
		case xltypeMulti:
			// This function does not accept array arguments.
			return &xlBadArgErr;
			break;
		case xltypeErr:
			// ToEvaluate is an error. Return Default instead.
			IsError = 1;
			break;
		default:
			return &xlBadArgErr;
			break;
	}

	if (IsError)
		return Default;
	else
		return ToEvaluate;
}

