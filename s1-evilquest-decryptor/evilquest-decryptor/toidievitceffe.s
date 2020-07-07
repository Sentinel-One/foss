//
//  toidievitceffe.s
//  evilquest-decryptor
//
//  Created by Julien-Pierre Avérous on 03/07/2020.
//


/*
** Extracted from EvilQuest / ThiefQuest ransomware
*/


// MARK: - Properties

.intel_syntax noprefix

// MARK: - Global symbols
.global _is_carved
.global __rotate;
.global _tpdcrypt;


// MARK: - Functions
.section __TEXT, __text

// MARK: > _rotate
//  - uint32_t _rotate(uint32_t a1, int8_t a2)
__rotate:
push       rbp
mov        rbp, rsp
mov        ax, si
mov        dword ptr [rbp-0x4], edi
mov        word ptr [rbp-0x6], ax
mov        esi, dword ptr [rbp-0x4]
movzx      ecx, word ptr [rbp-0x6]
shl        esi, cl
mov        edi, dword ptr [rbp-0x4]
movzx      edx, word ptr [rbp-0x6]
mov        r8d, edx
mov        r9d, 0x20
sub        r9, r8
mov        edx, r9d
mov        ecx, edx
shr        edi, cl
or         esi, edi
mov        eax, esi
pop        rbp
ret

// MARK: > is_carved
//  - bool is_carved(FILE *a1)
_is_carved:
push       rbp
mov        rbp, rsp
sub        rsp, 0x20
mov        qword ptr [rbp-0x8], rdi
mov        dword ptr [rbp-0xC], 0x0
mov        rdi, qword ptr [rbp-0x8]
mov        rsi, 0xfffffffffffffffc
mov        edx, 0x2
call       _fseek
lea        rsi, qword ptr [rbp-0xC]
mov        rcx, qword ptr [rbp-0x8]
mov        rdi, rsi
mov        esi, 0x1
mov        edx, 0x4
mov        dword ptr [rbp-0x10], eax
call       _fread
xor        r8d, r8d
mov        esi, r8d
xor        edx, edx
mov        rdi, qword ptr [rbp-0x8]
mov        qword ptr [rbp-0x18], rax
call       _fseek
cmp        dword ptr [rbp-0xC], 0xddbebabe
sete       r9b
and        r9b, 0x1
movzx      edx, r9b
mov        dword ptr [rbp-0x1C], eax
mov        eax, edx
add        rsp, 0x20
pop        rbp
ret

// MARK: > tpdcrypt
//  - int64_t tpdcrypt(const char *key, const void *encrypted_bytes, size_t encrypted_size, void **decrypted_bytes, size_t *descryted_size)
_tpdcrypt:
push       rbp
mov        rbp, rsp
sub        rsp, 0xe0
mov        rax, qword ptr [rip+L__stack_chk_guard]
mov        rax, qword ptr [rax]
mov        qword ptr [rbp-0x8], rax
mov        qword ptr [rbp-0x98], rdi
mov        qword ptr [rbp-0xA0], rsi
mov        qword ptr [rbp-0xA8], rdx
mov        qword ptr [rbp-0xB0], rcx
mov        qword ptr [rbp-0xB8], r8
mov        rax, qword ptr [rbp-0xA8]
sub        rax, 0x1
mov        qword ptr [rbp-0xA8], rax
mov        rax, qword ptr [rbp-0xA0]
mov        rcx, qword ptr [rbp-0xA8]
mov        r9b, byte ptr [rax+rcx]
mov        byte ptr [rbp-0xB9], r9b
mov        rax, qword ptr [rbp-0xA8]
movzx      r10d, byte ptr [rbp-0xB9]
mov        ecx, r10d
sub        rax, rcx
mov        rcx, qword ptr [rbp-0xB8]
mov        qword ptr [rcx], rax
mov        rsi, qword ptr [rbp-0xA8]
mov        edi, 0x1
call       _calloc
mov        rcx, 0xffffffffffffffff
mov        qword ptr [rbp-0xC8], rax
mov        rdi, qword ptr [rbp-0xC8]
mov        rsi, qword ptr [rbp-0xA0]
mov        rdx, qword ptr [rbp-0xA8]
call       ___memcpy_chk
mov        dword ptr [rbp-0xCC], 0x0
mov        dword ptr [rbp-0xD0], 0x2
mov        qword ptr [rbp-0xD8], rax

Lloc_100006e3a:
movsxd     rax, dword ptr [rbp-0xCC]
cmp        rax, qword ptr [rbp-0xA8]
jae        Lloc_100006ec5

lea        rdi, qword ptr [rbp-0x90]
mov        rsi, qword ptr [rbp-0x98]
mov        eax, dword ptr [rbp-0xD0]
mov        cl, al
mov        edx, 0x400
movzx      ecx, cl
call       __generate_xkey
lea        rdi, qword ptr [rbp-0x90]
mov        rsi, qword ptr [rbp-0xC8]
movsxd     r8, dword ptr [rbp-0xCC]
add        rsi, r8
mov        r8, qword ptr [rbp-0xA0]
movsxd     r9, dword ptr [rbp-0xCC]
add        r8, r9
mov        rdx, r8
call       __tp_decrypt
mov        eax, dword ptr [rbp-0xCC]
add        eax, 0x8
mov        dword ptr [rbp-0xCC], eax
mov        eax, dword ptr [rbp-0xD0]
add        eax, 0x1
mov        dword ptr [rbp-0xD0], eax
jmp        Lloc_100006e3a

Lloc_100006ec5:
mov        rdi, qword ptr [rbp-0xC8]
mov        rax, qword ptr [rbp-0xB8]
mov        rsi, qword ptr [rax]
call       _realloc
mov        qword ptr [rbp-0xC8], rax
mov        rax, qword ptr [rbp-0xC8]
mov        rsi, qword ptr [rbp-0xB0]
mov        qword ptr [rsi], rax
mov        rax, qword ptr [rip+L__stack_chk_guard]
mov        rax, qword ptr [rax]
mov        rsi, qword ptr [rbp-0x8]
cmp        rax, rsi
jne        Lloc_100006f13

add        rsp, 0xe0
pop        rbp
ret

Lloc_100006f13:
call       ___stack_chk_fail
ud2


// MARK: > _generate_xkey
//  - int64_t _generate_xkey(int64_t a1, char *a2, unsigned int a3, uint8_t a4)
__generate_xkey:
push       rbp
mov        rbp, rsp
sub        rsp, 0x160
mov        al, cl
mov        r8, qword ptr [rip+L__stack_chk_guard]
mov        r8, qword ptr [r8]
mov        qword ptr [rbp-0x8], r8
mov        qword ptr [rbp-0x118], rdi
mov        qword ptr [rbp-0x120], rsi
mov        dword ptr [rbp-0x124], edx
mov        byte ptr [rbp-0x125], al
mov        rdi, qword ptr [rbp-0x120]
call       _strlen
mov        ecx, eax
mov        dword ptr [rbp-0x130], ecx
lea        rax, qword ptr [rbp-0x110]
mov        rdi, rax
lea        rsi, qword ptr [rip+_bbbb+16]
mov        edx, 0x100
call       _memcpy
movzx      ecx, byte ptr [rbp-0x125]
mov        eax, ecx
cmp        rax, 0x20
jae        Lloc_10000624c

movzx      eax, byte ptr [rbp-0x125]
mov        ecx, eax
mov        qword ptr [rbp-0x140], rcx
jmp        Lloc_100006260

Lloc_10000624c:
movzx      eax, byte ptr [rbp-0x125]
mov        ecx, eax
and        rcx, 0x1f
mov        qword ptr [rbp-0x140], rcx

Lloc_100006260:
mov        rax, qword ptr [rbp-0x140]
lea        rcx, qword ptr [rip+_bbbb+272]
mov        dl, byte ptr [rcx+rax]
mov        byte ptr [rbp-0x125], dl
mov        word ptr [rbp-0x132], 0x0

Lloc_100006280:
movsx      eax, word ptr [rbp-0x132]
cmp        eax, 0x100
jge        Lloc_100006306

movsx      rax, word ptr [rbp-0x132]
movzx      ecx, byte ptr [rbp+rax-0x110]
movsx      rax, word ptr [rbp-0x132]
movzx      eax, byte ptr [rbp+rax-0x110]
movzx      edx, byte ptr [rbp-0x125]
mov        dword ptr [rbp-0x144], edx
cdq
mov        esi, dword ptr [rbp-0x144]
idiv       esi
add        ecx, edx
mov        eax, ecx
cdq
mov        ecx, 0xff
idiv       ecx
mov        dil, dl
movsx      r8, word ptr [rbp-0x132]
mov        byte ptr [rbp+r8-0x110], dil
movzx      eax, byte ptr [rbp-0x125]
movsx      ecx, word ptr [rbp-0x132]
add        ecx, eax
mov        dx, cx
mov        word ptr [rbp-0x132], dx
jmp        Lloc_100006280

Lloc_100006306:
cmp        dword ptr [rbp-0x130], 0x0
jne        Lloc_100006318

jmp        Lloc_1000065e6

Lloc_100006318:
cmp        dword ptr [rbp-0x130], 0x80
jbe        Lloc_10000634a

mov        rdi, qword ptr [rbp-0x120]
mov        esi, 0x80
call       _realloc
mov        qword ptr [rbp-0x120], rax
mov        dword ptr [rbp-0x130], 0x80

Lloc_10000634a:
cmp        dword ptr [rbp-0x124], 0x400
setbe      al
xor        al, 0xff
and        al, 0x1
movzx      ecx, al
movsxd     rdx, ecx
cmp        rdx, 0x0
je         Lloc_10000638a

lea        rdi, qword ptr [rip+aGeneratexkey]
lea        rsi, qword ptr [rip+aToidievitceffe_100011081]
lea        rcx, qword ptr [rip+aBits1024]
mov        edx, 0x57
call       ___assert_rtn

Lloc_10000638a:
jmp        Lloc_10000638f

Lloc_10000638f:
cmp        dword ptr [rbp-0x124], 0x0
jne        Lloc_1000063a6

mov        dword ptr [rbp-0x124], 0x400

Lloc_1000063a6:
mov        rcx, 0xffffffffffffffff
mov        rax, qword ptr [rbp-0x118]
mov        rsi, qword ptr [rbp-0x120]
mov        edx, dword ptr [rbp-0x130]
mov        rdi, rax
call       ___memcpy_chk
cmp        dword ptr [rbp-0x130], 0x80
mov        qword ptr [rbp-0x150], rax
jae        Lloc_100006485

mov        dword ptr [rbp-0x12C], 0x0
mov        rax, qword ptr [rbp-0x118]
mov        ecx, dword ptr [rbp-0x130]
sub        ecx, 0x1
mov        ecx, ecx
mov        edx, ecx
mov        sil, byte ptr [rax+rdx]
mov        byte ptr [rbp-0x126], sil

Lloc_100006409:
movzx      eax, byte ptr [rbp-0x126]
mov        rcx, qword ptr [rbp-0x118]
mov        edx, dword ptr [rbp-0x12C]
mov        esi, edx
add        esi, 0x1
mov        dword ptr [rbp-0x12C], esi
mov        edx, edx
mov        edi, edx
movzx      edx, byte ptr [rcx+rdi]
add        eax, edx
and        eax, 0xff
movsxd     rcx, eax
mov        r8b, byte ptr [rbp+rcx-0x110]
mov        byte ptr [rbp-0x126], r8b
mov        r8b, byte ptr [rbp-0x126]
mov        rcx, qword ptr [rbp-0x118]
mov        eax, dword ptr [rbp-0x130]
mov        edx, eax
add        edx, 0x1
mov        dword ptr [rbp-0x130], edx
mov        eax, eax
mov        edi, eax
mov        byte ptr [rcx+rdi], r8b
cmp        dword ptr [rbp-0x130], 0x80
jb         Lloc_100006409

jmp        Lloc_100006485

Lloc_100006485:
xor        eax, eax
mov        ecx, dword ptr [rbp-0x124]
add        ecx, 0x7
shr        ecx, 0x3
mov        dword ptr [rbp-0x130], ecx
mov        ecx, 0x80
sub        ecx, dword ptr [rbp-0x130]
mov        dword ptr [rbp-0x12C], ecx
mov        rdx, qword ptr [rbp-0x118]
mov        ecx, dword ptr [rbp-0x12C]
mov        esi, ecx
movzx      ecx, byte ptr [rdx+rsi]
sub        eax, dword ptr [rbp-0x124]
and        eax, 0x7
mov        dword ptr [rbp-0x154], ecx
mov        ecx, eax
mov        eax, 0xff
sar        eax, cl
mov        edi, dword ptr [rbp-0x154]
and        edi, eax
movsxd     rdx, edi
mov        cl, byte ptr [rbp+rdx-0x110]
mov        byte ptr [rbp-0x126], cl
mov        cl, byte ptr [rbp-0x126]
mov        rdx, qword ptr [rbp-0x118]
mov        eax, dword ptr [rbp-0x12C]
mov        esi, eax
mov        byte ptr [rdx+rsi], cl

Lloc_100006505:
mov        eax, dword ptr [rbp-0x12C]
mov        ecx, eax
add        ecx, 0xffffffff
mov        dword ptr [rbp-0x12C], ecx
cmp        eax, 0x0
je         Lloc_100006574

movzx      eax, byte ptr [rbp-0x126]
mov        rcx, qword ptr [rbp-0x118]
mov        edx, dword ptr [rbp-0x12C]
add        edx, dword ptr [rbp-0x130]
mov        edx, edx
mov        esi, edx
movzx      edx, byte ptr [rcx+rsi]
xor        eax, edx
movsxd     rcx, eax
mov        dil, byte ptr [rbp+rcx-0x110]
mov        byte ptr [rbp-0x126], dil
mov        dil, byte ptr [rbp-0x126]
mov        rcx, qword ptr [rbp-0x118]
mov        eax, dword ptr [rbp-0x12C]
mov        esi, eax
mov        byte ptr [rcx+rsi], dil
jmp        Lloc_100006505

Lloc_100006574:
mov        dword ptr [rbp-0x12C], 0x3f

Lloc_10000657e:
mov        rax, qword ptr [rbp-0x118]
mov        ecx, dword ptr [rbp-0x12C]
shl        ecx, 0x1
mov        ecx, ecx
mov        edx, ecx
movzx      ecx, byte ptr [rax+rdx]
mov        rax, qword ptr [rbp-0x118]
mov        esi, dword ptr [rbp-0x12C]
shl        esi, 0x1
add        esi, 0x1
mov        esi, esi
mov        edx, esi
movzx      esi, byte ptr [rax+rdx]
shl        esi, 0x8
add        ecx, esi
mov        di, cx
mov        rax, qword ptr [rbp-0x118]
mov        ecx, dword ptr [rbp-0x12C]
mov        edx, ecx
mov        word ptr [rax+rdx*2], di
mov        eax, dword ptr [rbp-0x12C]
mov        ecx, eax
add        ecx, 0xffffffff
mov        dword ptr [rbp-0x12C], ecx
cmp        eax, 0x0
jne        Lloc_10000657e

Lloc_1000065e6:
mov        rax, qword ptr [rip+L__stack_chk_guard]
mov        rax, qword ptr [rax]
mov        rcx, qword ptr [rbp-0x8]
cmp        rax, rcx
jne        Lloc_100006606

add        rsp, 0x160
pop        rbp
ret

Lloc_100006606:
call       ___stack_chk_fail


// MARK: > _tp_decrypt
// int64 _tp_decrypt(int64 a1, int64 a2, uint8 *a3)
__tp_decrypt:
push       rbp
mov        rbp, rsp
mov        qword ptr [rbp-0x8], rdi
mov        qword ptr [rbp-0x10], rsi
mov        qword ptr [rbp-0x18], rdx
mov        rdx, qword ptr [rbp-0x18]
movzx      eax, byte ptr [rdx+7]
shl        eax, 0x8
mov        rdx, qword ptr [rbp-0x18]
movzx      ecx, byte ptr [rdx+6]
add        eax, ecx
mov        dword ptr [rbp-0x1C], eax
mov        rdx, qword ptr [rbp-0x18]
movzx      eax, byte ptr [rdx+5]
shl        eax, 0x8
mov        rdx, qword ptr [rbp-0x18]
movzx      ecx, byte ptr [rdx+4]
add        eax, ecx
mov        dword ptr [rbp-0x20], eax
mov        rdx, qword ptr [rbp-0x18]
movzx      eax, byte ptr [rdx+3]
shl        eax, 0x8
mov        rdx, qword ptr [rbp-0x18]
movzx      ecx, byte ptr [rdx+2]
add        eax, ecx
mov        dword ptr [rbp-0x24], eax
mov        rdx, qword ptr [rbp-0x18]
movzx      eax, byte ptr [rdx+1]
shl        eax, 0x8
mov        rdx, qword ptr [rbp-0x18]
movzx      ecx, byte ptr [rdx]
add        eax, ecx
mov        dword ptr [rbp-0x28], eax
mov        dword ptr [rbp-0x2C], 0xf

Lloc_100006916:
mov        eax, dword ptr [rbp-0x1C]
and        eax, 0xffff
mov        dword ptr [rbp-0x1C], eax
mov        eax, dword ptr [rbp-0x1C]
shl        eax, 0xb
mov        ecx, dword ptr [rbp-0x1C]
shr        ecx, 0x5
add        eax, ecx
mov        dword ptr [rbp-0x1C], eax
mov        eax, dword ptr [rbp-0x28]
mov        ecx, dword ptr [rbp-0x20]
xor        ecx, 0xffffffff
and        eax, ecx
mov        ecx, dword ptr [rbp-0x24]
and        ecx, dword ptr [rbp-0x20]
add        eax, ecx
mov        rdx, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x2C]
shl        ecx, 0x2
add        ecx, 0x3
mov        ecx, ecx
mov        esi, ecx
movzx      ecx, word ptr [rdx+rsi*2]
add        eax, ecx
mov        ecx, dword ptr [rbp-0x1C]
sub        ecx, eax
mov        dword ptr [rbp-0x1C], ecx
mov        eax, dword ptr [rbp-0x20]
and        eax, 0xffff
mov        dword ptr [rbp-0x20], eax
mov        eax, dword ptr [rbp-0x20]
shl        eax, 0xd
mov        ecx, dword ptr [rbp-0x20]
shr        ecx, 0x3
add        eax, ecx
mov        dword ptr [rbp-0x20], eax
mov        eax, dword ptr [rbp-0x1C]
mov        ecx, dword ptr [rbp-0x24]
xor        ecx, 0xffffffff
and        eax, ecx
mov        ecx, dword ptr [rbp-0x28]
and        ecx, dword ptr [rbp-0x24]
add        eax, ecx
mov        rdx, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x2C]
shl        ecx, 0x2
add        ecx, 0x2
mov        ecx, ecx
mov        esi, ecx
movzx      ecx, word ptr [rdx+rsi*2]
add        eax, ecx
mov        ecx, dword ptr [rbp-0x20]
sub        ecx, eax
mov        dword ptr [rbp-0x20], ecx
mov        eax, dword ptr [rbp-0x24]
and        eax, 0xffff
mov        dword ptr [rbp-0x24], eax
mov        eax, dword ptr [rbp-0x24]
shl        eax, 0xe
mov        ecx, dword ptr [rbp-0x24]
shr        ecx, 0x2
add        eax, ecx
mov        dword ptr [rbp-0x24], eax
mov        eax, dword ptr [rbp-0x20]
mov        ecx, dword ptr [rbp-0x28]
xor        ecx, 0xffffffff
and        eax, ecx
mov        ecx, dword ptr [rbp-0x1C]
and        ecx, dword ptr [rbp-0x28]
add        eax, ecx
mov        rdx, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x2C]
shl        ecx, 0x2
add        ecx, 0x1
mov        ecx, ecx
mov        esi, ecx
movzx      ecx, word ptr [rdx+rsi*2]
add        eax, ecx
mov        ecx, dword ptr [rbp-0x24]
sub        ecx, eax
mov        dword ptr [rbp-0x24], ecx
mov        eax, dword ptr [rbp-0x28]
and        eax, 0xffff
mov        dword ptr [rbp-0x28], eax
mov        eax, dword ptr [rbp-0x28]
shl        eax, 0xf
mov        ecx, dword ptr [rbp-0x28]
shr        ecx, 0x1
add        eax, ecx
mov        dword ptr [rbp-0x28], eax
mov        eax, dword ptr [rbp-0x24]
mov        ecx, dword ptr [rbp-0x1C]
xor        ecx, 0xffffffff
and        eax, ecx
mov        ecx, dword ptr [rbp-0x20]
and        ecx, dword ptr [rbp-0x1C]
add        eax, ecx
mov        rdx, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x2C]
shl        ecx, 0x2
add        ecx, 0x0
mov        ecx, ecx
mov        esi, ecx
movzx      ecx, word ptr [rdx+rsi*2]
add        eax, ecx
mov        ecx, dword ptr [rbp-0x28]
sub        ecx, eax
mov        dword ptr [rbp-0x28], ecx
cmp        dword ptr [rbp-0x2C], 0x5
je         Lloc_100006a62

cmp        dword ptr [rbp-0x2C], 0xb
jne        Lloc_100006aca

Lloc_100006a62:
mov        rax, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x20]
and        ecx, 0x3f
mov        ecx, ecx
mov        edx, ecx
movzx      ecx, word ptr [rax+rdx*2]
mov        esi, dword ptr [rbp-0x1C]
sub        esi, ecx
mov        dword ptr [rbp-0x1C], esi
mov        rax, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x24]
and        ecx, 0x3f
mov        ecx, ecx
mov        edx, ecx
movzx      ecx, word ptr [rax+rdx*2]
mov        esi, dword ptr [rbp-0x20]
sub        esi, ecx
mov        dword ptr [rbp-0x20], esi
mov        rax, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x28]
and        ecx, 0x3f
mov        ecx, ecx
mov        edx, ecx
movzx      ecx, word ptr [rax+rdx*2]
mov        esi, dword ptr [rbp-0x24]
sub        esi, ecx
mov        dword ptr [rbp-0x24], esi
mov        rax, qword ptr [rbp-0x8]
mov        ecx, dword ptr [rbp-0x1C]
and        ecx, 0x3f
mov        ecx, ecx
mov        edx, ecx
movzx      ecx, word ptr [rax+rdx*2]
mov        esi, dword ptr [rbp-0x28]
sub        esi, ecx
mov        dword ptr [rbp-0x28], esi

Lloc_100006aca:
jmp        Lloc_100006acf

Lloc_100006acf:
mov        eax, dword ptr [rbp-0x2C]
mov        ecx, eax
add        ecx, 0xffffffff
mov        dword ptr [rbp-0x2C], ecx
cmp        eax, 0x0
jne        Lloc_100006916

mov        eax, dword ptr [rbp-0x28]
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx], cl
mov        eax, dword ptr [rbp-0x28]
shr        eax, 0x8
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+1], cl
mov        eax, dword ptr [rbp-0x24]
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+2], cl
mov        eax, dword ptr [rbp-0x24]
shr        eax, 0x8
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+3], cl
mov        eax, dword ptr [rbp-0x20]
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+4], cl
mov        eax, dword ptr [rbp-0x20]
shr        eax, 0x8
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+5], cl
mov        eax, dword ptr [rbp-0x1C]
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+6], cl
mov        eax, dword ptr [rbp-0x1C]
shr        eax, 0x8
mov        cl, al
mov        rdx, qword ptr [rbp-0x10]
mov        byte ptr [rdx+7], cl
pop        rbp
ret


// MARK: - Resources
// MARK:  > cstring
.section __TEXT, __cstring

aGeneratexkey:
.string "_generate_xkey"

aToidievitceffe_100011081:
.string "/toidievitceffe/libtpyrc/tpyrc.c"

aBits1024:
.string "bits <= 1024"


// MARK:  > data
.section __DATA, __data

_bbbb:
.byte 0x67, 0x45, 0x8B, 0x6B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xD9, 0x78, 0xF9, 0xC4, 0x19, 0xDD, 0xB5
.byte 0xED, 0x28, 0xE9, 0xFD, 0x79, 0x4A, 0xA0, 0xD8, 0x9D, 0xC6, 0x7E, 0x37, 0x83, 0x2B, 0x76, 0x53, 0x8E, 0x62, 0x4C, 0x64, 0x88, 0x44, 0x8B
.byte 0xFB, 0xA2, 0x17, 0x9A, 0x59, 0xF5, 0x87, 0xB3, 0x4F, 0x13, 0x61, 0x45, 0x6D, 0x8D, 0x09, 0x81, 0x7D, 0x32, 0xBD, 0xC9, 0x40, 0xEB, 0x86
.byte 0xB7, 0x7B, 0x0B, 0xF0, 0x95, 0x21, 0x22, 0x5C, 0x6B, 0x4E, 0x82, 0x54, 0xD6, 0x65, 0x93, 0xCE, 0x60, 0xB2, 0x1C, 0x73, 0x56, 0x71, 0x14
.byte 0xA7, 0x8C, 0xF1, 0xDC, 0x12, 0x75, 0xCA, 0x1F, 0x3B, 0xBE, 0xE4, 0xD1, 0x42, 0x3D, 0xD4, 0x30, 0xA3, 0x3C, 0xB6, 0x26, 0x6F, 0xBF, 0x0E
.byte 0xDA, 0x46, 0x69, 0x07, 0x57, 0x27, 0xF2, 0xD2, 0x9B, 0xBC, 0x94, 0x43, 0x03, 0xF8, 0x11, 0x6C, 0xF6, 0x90, 0xEF, 0x3E, 0xE7, 0x06, 0xC3
.byte 0xD5, 0x2F, 0xC8, 0x66, 0x1E, 0xD7, 0x08, 0xE8, 0xEA, 0xDE, 0x80, 0x52, 0xEE, 0xF7, 0x84, 0xAA, 0x72, 0xAC, 0x35, 0x4D, 0x6A, 0x2A, 0x96
.byte 0x1A, 0x1D, 0xC0, 0x5A, 0x15, 0x49, 0x74, 0x4B, 0x9F, 0xD0, 0x5E, 0x04, 0x18, 0xA4, 0xEC, 0xC2, 0xE0, 0x41, 0x6E, 0x0F, 0x51, 0xCB, 0xCC
.byte 0x24, 0x91, 0xAF, 0x50, 0xA1, 0xF4, 0x70, 0x39, 0x99, 0x7C, 0x3A, 0x85, 0x23, 0xB8, 0xB4, 0x7A, 0xFC, 0x02, 0x36, 0x5B, 0x25, 0x55, 0x97
.byte 0x31, 0x2D, 0x5D, 0xFA, 0x98, 0xE3, 0x8A, 0x92, 0xAE, 0x05, 0xDF, 0x29, 0x10, 0x67, 0xC7, 0xBA, 0x8F, 0xD3, 0x00, 0xE6, 0xCF, 0xE1, 0x9E
.byte 0xA8, 0x2C, 0x63, 0x16, 0x01, 0x3F, 0x58, 0xE2, 0x89, 0xA9, 0x0D, 0x38, 0x34, 0x1B, 0xAB, 0x33, 0xFF, 0xB0, 0xBB, 0x7F, 0x0C, 0x5F, 0xB9
.byte 0xB1, 0xCD, 0x2E, 0xC5, 0xF3, 0xDB, 0x47, 0xE5, 0xA5, 0x9C, 0x77, 0x0A, 0xA6, 0x20, 0x68, 0xFE, 0x48, 0xC1, 0xAD, 0x02, 0x03, 0x05, 0x07
.byte 0x0B, 0x0D, 0x11, 0x13, 0x17, 0x1D, 0x1F, 0x25, 0x29, 0x2B, 0x2F, 0x35, 0x3B, 0x3D, 0x43, 0x47, 0x49, 0x4F, 0x53, 0x59, 0x61, 0x65, 0x67
.byte 0x6B, 0x6D, 0x71, 0x7F, 0x83, 0x01, 0x00, 0x00, 0x00, 0x0E, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00


// MARK:  > got
.section __DATA, __got

L__stack_chk_guard:
.quad  ___stack_chk_guard
