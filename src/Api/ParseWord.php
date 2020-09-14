<?php
namespace MDword\Api;

class ParseWord extends Base{
    public function getBlockList() {
        $blocks = $this->wordProcessor->getBlockList();
        $this->success($blocks);
    }
}
