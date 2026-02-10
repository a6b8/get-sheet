import { join } from 'node:path'


const PROJECT_CWD = join( import.meta.dirname, '..', '..' )

const VALID_CREDENTIALS_PATH = '/tmp/test-credentials.json'

const VALID_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'

const VALID_TAB = 'Sheet1'

const VALID_DATA_2D = [
    [ 'Name', 'Score' ],
    [ 'Alice', 95 ],
    [ 'Bob', 87 ]
]


export {
    PROJECT_CWD,
    VALID_CREDENTIALS_PATH,
    VALID_SPREADSHEET_ID,
    VALID_TAB,
    VALID_DATA_2D
}
