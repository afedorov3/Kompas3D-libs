#ifndef _COMMON_H
#define _COMMON_H

// common errors
enum {
	LIBSTATUS_SUCCESS	= 0,
	// exclusive status, bit 29 can only be used by LIBSTATUS_SYSERR
	LIBSTATUS_ERR		= 0x10000000, // exclusive errors
	LIBSTATUS_SYSERR	= 0x20000000, // exclusive system errors, MSFT: "Bit 29 is reserved for application-defined error codes; no system error code has this bit set."
	LIBSTATUS_WARN		= 0x40000000, // exclusive warnings
	LIBSTATUS_INFO		= 0x50000000, // exclusive informational
	LIBSTATUS_EXCLUSIVE	= 0xF0000000, // exclusive mask

	LIBSTATUS_ERR_ABORTED				= LIBSTATUS_ERR | 0xC000000,
	LIBSTATUS_ERR_COMMON_MIN			= LIBSTATUS_ERR_ABORTED,
	LIBSTATUS_ERR_API					= LIBSTATUS_ERR | 0xD000000,
	LIBSTATUS_ERR_INVARG				= LIBSTATUS_ERR | 0xE000000,
	LIBSTATUS_ERR_UNKNOWN				= LIBSTATUS_ERR | 0xF000000,

	LIBSTATUS_SYSERR_NO_MEM				= LIBSTATUS_SYSERR | ERROR_NOT_ENOUGH_MEMORY,
};

#endif /* _COMMON_H */
