export class ComError extends Error {
  constructor(
    message: string,
    public readonly recoverable: boolean = false,
  ) {
    super(message)
    this.name = "ComError"
  }
}

export class TransientComError extends ComError {
  constructor(message: string) {
    super(message, true)
    this.name = "TransientComError"
  }
}

export class FatalComError extends ComError {
  constructor(message: string) {
    super(message, false)
    this.name = "FatalComError"
  }
}
