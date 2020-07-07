/*
 * main.m
 * evilquest-decryptor
 *
 * Created by Julien-Pierre Avérous on 03/07/2020.
 * Copyright (c) 2020 SentinelOne. Licensed under the MIT License.
 *
 *
 * ## MIT License ##
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */


#import <Foundation/Foundation.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <errno.h>


/*
** EvilQuest Imports
*/
#pragma mark - EvilQuest Imports

extern BOOL		is_carved(FILE *file);
extern uint32_t _rotate(uint32_t a1, char a2);

int64_t tpdcrypt(const char *key, const void *encrypted_bytes, size_t encrypted_size, void **decrypted_bytes, size_t *descryted_size);



/*
** Main
*/
#pragma mark - Main

int main(int argc, const char * argv[])
{
	// Banner.
	NSDictionary	*plist = [[NSBundle mainBundle] infoDictionary];
	NSString		*version = plist[@"CFBundleShortVersionString"];
	NSString		*build = plist[@"CFBundleVersion"];

	fprintf(stderr, " ~~~  SentinelOne's EvilQuest decryptor v.%s (%s) ~~~\n", version.UTF8String, build.UTF8String);
	fprintf(stderr, "  by:\n");
	fprintf(stderr, "   - Julien-Pierre Avérous\n");
	fprintf(stderr, "   - Jason Reaves\n");
	fprintf(stderr, "   - Phil Stokes\n");
	fprintf(stderr, "   - Jim Walter\n");
	fprintf(stderr, "\n");

	// Check & fetch parameters.
	if (argc != 3)
	{
		fprintf(stderr, "Usage: %s <input-encrypted-file> <output-decrypted-file>\n", getprogname());
		return 1;
	}

	const char *inputPath = argv[1];
	const char *outputPath = argv[2];

	// Open input file.
	FILE *input = fopen(inputPath, "rb");

	if (!input)
	{
		fprintf(stderr, "Cannot open '%s': %s\n", inputPath, strerror(errno));
		return 1;
	}

	// Check if file is carved.
	if (!is_carved(input))
	{
		fprintf(stderr, "File '%s' is not encrypted.\n", inputPath);
		return 1;
	}

	// Open output file.
	FILE *output = fopen(outputPath, "wb");

	if (!output)
	{
		fprintf(stderr, "Cannot open '%s': %s\n", outputPath, strerror(errno));
		return 1;
	}

	// Read entries.
	// > Metadata.
	uint32_t metaLength = 0;

	fseek(input, -12, SEEK_END);

	if (fread(&metaLength, sizeof(metaLength), 1, input) != 1)
	{
		fprintf(stderr, "Failed to read metadatas.\n");
		return 1;
	}

	// > Key size.
	uint32_t keySize = 0;

	fseek(input, -(long)metaLength - 12, SEEK_END);

	if (fread(&keySize, sizeof(keySize), 1, input) != 1)
	{
		fprintf(stderr, "Failed to read key-size.\n");
		return 1;
	}

	if (keySize != 0x80)
	{
		fprintf(stderr, "Invalid key size: %x.\n", keySize);
		return 1;
	}

	// > Block size.
	uint32_t blockSize = 0;

	fseek(input, -(long)metaLength - 20, SEEK_END);

	if (fread(&blockSize, sizeof(blockSize), 1, input) != 1)
	{
		fprintf(stderr, "Failed to read block-size.\n");
		return 1;
	}

	// > Encrypted key.
	uint8_t encryptedKey[132] = { 0 };

	fseek(input, -(long)metaLength + 8, SEEK_END);

	if (fread(encryptedKey, sizeof(encryptedKey), 1, input) != 1)
	{
		fprintf(stderr, "Failed to read encrypted key.\n");
		return 1;
	}

	// Extract seed.
	uint32_t *seedPtr = (uint32_t *)(encryptedKey + sizeof(encryptedKey) - sizeof(uint32_t));
	uint32_t seed = *seedPtr;

	// Extract key.
	// > Decrypt key.
	NSMutableData 	*keyData = [NSMutableData data];
	uint32_t		*blocks = (uint32_t *)encryptedKey;

	for (unsigned i = 0; i < sizeof(encryptedKey) / sizeof(*blocks); i++)
	{
		uint32_t value = blocks[i] ^ seed;

		[keyData appendBytes:&value length:sizeof(value)];

		seed = _rotate(seed, 1);
	}

	[keyData setLength:0x80];

	// > Convert & check.
	NSString *keyStr = [[NSString alloc] initWithData:keyData encoding:NSASCIIStringEncoding];

	if ([keyStr rangeOfCharacterFromSet:[[NSCharacterSet alphanumericCharacterSet] invertedSet]].location != NSNotFound)
	{
		fprintf(stderr, "Encryption key is invalid.\n");
		return 1;
	}

	// > Print.
	fprintf(stderr, "The encryption key of this file is '%s'.\n", keyStr.UTF8String);

	// Compute file length.
	size_t inputLength;

	fseek(input, 0, SEEK_END);
	inputLength = ftell(input);
	fseek(input, 0, SEEK_SET);

	// Compute input length.
	inputLength -= metaLength + 20;

	// Decrypt file.
	uint8_t	readBytes[blockSize];

	while (inputLength > 0)
	{
		// > Read a block of data.
		size_t readSize = MIN(blockSize, inputLength);
		size_t readResult = fread(readBytes, readSize, 1, input);

		if (readResult != 1)
		{
			fprintf(stderr, "Failed to read input file.\n");
			return 1;
		}

		inputLength -= readSize;

		// > Decrypt the block of data.
		void	*outputPtr = NULL;
		size_t	outputSize = 0;

		tpdcrypt(keyStr.UTF8String, readBytes, readSize, &outputPtr, &outputSize);

		if (!outputPtr || outputSize == 0)
		{
			fprintf(stderr, "Failed to decrypt block of data.\n");
			return 1;
		}

		// > Write decrypted block of data.
		if (fwrite(outputPtr, outputSize, 1, output) != 1)
		{
			fprintf(stderr, "Failed to write output file.\n");
			return 1;
		}

		// > Clean.
		free(outputPtr);
	}

	fprintf(stderr, "Your file was decrypted with success: '%s'\n", outputPath);

	return 0;
}
