// The one error every encrypted format raises when its password does not open
// it — an Office package and a PDF alike, so that a caller asks for a password
// with one `instanceof` whatever it was handed.

/**
 * Thrown when the bytes are encrypted and the password does not open them —
 * the wrong one, or none where one is needed.
 */
export class WrongPasswordError extends Error {
  /** @param message The message to carry; the default states the failure. */
  constructor(message = 'Wrong password for this document') {
    super(message);
    this.name = 'WrongPasswordError';
  }
}
