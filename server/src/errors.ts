export class AppError extends Error {
    code: string; http?: number; extra?: any;
    constructor(code: string, message: string, http = 400, extra?: any) {
        super(message); this.code = code; this.http = http; this.extra = extra;
    }
}
export const ERR = {
    AUTH: (msg = 'unauthorized') => new AppError('E_AUTH', msg, 401),
    VALID: (msg = 'invalid input', extra?: any) => new AppError('E_VALID', msg, 400, extra),
    NOTFOUND: (msg = 'not found') => new AppError('E_NOTFOUND', msg, 404),
    CONFLICT: (msg = 'conflict', extra?: any) => new AppError('E_CONFLICT', msg, 409, extra),
};
