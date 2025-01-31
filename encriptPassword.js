function hashPasswordAndLog(password) {
    // Hashear la contraseña
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password.trim());

    // Convertir el hash a formato hexadecimal
    const hashedPassword = digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');

    // Mostrar la contraseña hasheada en la consola
    console.log(`Contraseña hasheada: ${hashedPassword}`);
}

// Ejemplo de uso
function testHashPassword() {
    hashPasswordAndLog('chavezs2025$');
}
