<?php

declare(strict_types=1);

namespace Slam\Excel\Helper;

use ArrayAccess;

/**
 * @extends ArrayAccess<string, ColumnInterface>
 */
interface ColumnCollectionInterface extends ArrayAccess
{
}
