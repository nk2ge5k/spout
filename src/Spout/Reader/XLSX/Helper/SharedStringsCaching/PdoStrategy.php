<?php

namespace Box\Spout\Reader\XLSX\Helper\SharedStringsCaching;

class PdoStrategy implements CachingStrategyInterface
{
    /**
     * @var \PDO
     */
    protected $pdo = null;

    protected $hash;

    public function __construct($pdo)
    {
        $this->pdo = $pdo;
        $this->hash = \bin2hex(\mcrypt_create_iv(22, MCRYPT_DEV_URANDOM));
        $this->createCacheTable();
    }

    /**
     * @inheritDoc
     */
    public function addStringForIndex($sharedString, $sharedStringIndex)
    {
        $sql = 'REPLACE INTO cache (key, value, hash) VALUES (?, ?, ?)';
        $stmt = $this->pdo->prepare($sql);

        if(false == $stmt) {
            throw new \Exception(implode(' ', $this->pdo->errorInfo()));
        }

        $params = [
            $sharedStringIndex, $sharedString, $this->hash
        ];

        if ($stmt->execute($params) === false) {
            throw new \Exception(
                "Caching $sharedStringIndex failed: ".implode(' ', $stmt->errorInfo())
            );
        }

        return true;
    }

    /**
     * @inheritDoc
     */
    public function closeCache()
    {
        return true;
    }

    /**
     * @inheritDoc
     */
    public function getStringAtIndex($sharedStringIndex)
    {
        $sql = 'SELECT value FROM cache WHERE key=? AND hash=?';
        $stmt = $this->pdo->prepare($sql);
        if( $stmt->execute([$sharedStringIndex, $this->hash]) === false) {
            throw new \Exception(
                "Fetching $sharedStringIndex failed: ".implode(' ', $stmt->errorInfo())
            );
        }
        $row = $stmt->fetch(\PDO::FETCH_ASSOC);

        $value = isset($row['value']) ? $row['value'] : null;

        return $value;

    }

    /**
     * @inheritDoc
     */
    public function clearCache()
    {
        return true;
    }

    protected function createCacheTable()
    {
        $sql = 'CREATE TABLE IF NOT EXISTS cache (id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, key TEXT, value TEXT, hash TEXT, UNIQUE (hash))';

        $result = $this->pdo->exec($sql);
        if ($result === false) {
            throw new \Exception('Could not create table: ' . implode(' ', $this->pdo->errorInfo()));
        }
    }

}