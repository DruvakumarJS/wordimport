<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Quiz extends Model
{
    use HasFactory;

    protected $fillable=[
            'question',
            'option_a',
            'option_b',
            'option_c',
            'option_d',
            'correct_option'
        ];
}
