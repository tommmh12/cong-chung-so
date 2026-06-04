const { Pool } = require('pg');
require('dotenv').config();

let connectionConfig = {};

if (process.env.DATABASE_URL) {
  connectionConfig = {
    connectionString: process.env.DATABASE_URL
  };
} else if (process.env.SUPABASE_DB_HOST && process.env.SUPABASE_DB_PASSWORD) {
  connectionConfig = {
    host: process.env.SUPABASE_DB_HOST,
    port: Number(process.env.SUPABASE_DB_PORT || 5432),
    database: process.env.SUPABASE_DB_NAME || 'postgres',
    user: process.env.SUPABASE_DB_USER || 'postgres',
    password: process.env.SUPABASE_DB_PASSWORD
  };
} else {
  throw new Error(
    'Missing database config. Set DATABASE_URL or SUPABASE_DB_HOST/SUPABASE_DB_PASSWORD in backend/.env'
  );
}

const rawPool = new Pool({
  ...connectionConfig,
  max: 10,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000,
  ssl: {
    rejectUnauthorized: false
  }
});

function toPgSql(sql) {
  let index = 0;
  return sql.replace(/\?/g, () => {
    index += 1;
    return `$${index}`;
  });
}

async function runQuery(executor, sql, params = []) {
  const result = await executor.query(toPgSql(sql), params);
  return [
    result.rows,
    {
      affectedRows: result.rowCount,
      rowCount: result.rowCount
    }
  ];
}

function wrapClient(client) {
  return {
    query(sql, params = []) {
      return runQuery(client, sql, params);
    },
    async beginTransaction() {
      await client.query('BEGIN');
    },
    async commit() {
      await client.query('COMMIT');
    },
    async rollback() {
      await client.query('ROLLBACK');
    },
    release() {
      client.release();
    }
  };
}

const pool = {
  query(sql, params = []) {
    return runQuery(rawPool, sql, params);
  },
  async getConnection() {
    const client = await rawPool.connect();
    return wrapClient(client);
  },
  end() {
    return rawPool.end();
  }
};

async function testConnection() {
  let client;

  try {
    client = await rawPool.connect();
    const result = await client.query('SELECT current_database(), current_user, version()');
    console.log('Connected to Supabase PostgreSQL successfully!');
    console.log('Details:', result.rows[0]);
  } catch (error) {
    console.error('Supabase connection failed:', error.message || error);
    if (error.cause) {
      console.error('Connection cause:', error.cause.message || error.cause);
    }
  } finally {
    client?.release();
  }
}

module.exports = {
  pool,
  testConnection
};
