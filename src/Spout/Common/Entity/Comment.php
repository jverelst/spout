<?php

namespace Box\Spout\Common\Entity;

class Comment
{
    /**
     * The value of this comment
     * @var mixed|null
     */
    protected $value = null;

    /**
     * @param mixed|null $value
     */
    public function setValue($value)
    {
        $this->value = $value;
    }

    /**
     * @return mixed|null
     */
    public function getValue()
    {
        return $this->value;
    }
}