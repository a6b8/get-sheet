export default {
    testEnvironment: 'node',
    transform: {},
    verbose: true,
    roots: ['./tests'],
    collectCoverageFrom: [
        'src/**/*.mjs',
        '!src/data/**',
        '!**/node_modules/**'
    ],
    coverageThreshold: {
        global: {
            branches: 15,
            functions: 15,
            lines: 15,
            statements: 15
        }
    },
    coverageDirectory: 'coverage',
    coverageReporters: ['text', 'lcov', 'html', 'json'],
    clearMocks: true,
    testTimeout: 10000
}
