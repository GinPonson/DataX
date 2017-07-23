package com.alibaba.datax.plugin.reader.filereader;

import com.alibaba.datax.common.spi.ErrorCode;

/**
 * Created by haiwei.luo on 14-9-20.
 */
public enum FileReaderErrorCode implements ErrorCode {
	REQUIRED_VALUE("FileReader-00", "您缺失了必须填写的参数值."),
	ILLEGAL_VALUE("FileReader-01", "您填写的参数值不合法."),
	MIXED_INDEX_VALUE("FileReader-02", "您的列信息配置同时包含了index,value."),
	NO_INDEX_VALUE("FileReader-03","您明确的配置列信息,但未填写相应的index,value."),
	FILE_NOT_EXISTS("FileReader-04", "您配置的目录文件路径不存在."),
	OPEN_FILE_WITH_CHARSET_ERROR("FileReader-05", "您配置的文件编码和实际文件编码不符合."),
	OPEN_FILE_ERROR("FileReader-06", "您配置的文件在打开时异常,建议您检查源目录是否有隐藏文件,管道文件等特殊文件."),
	READ_FILE_IO_ERROR("FileReader-07", "您配置的文件在读取时出现IO异常."),
	SECURITY_NOT_ENOUGH("FileReader-08", "您缺少权限执行相应的文件操作."),
	CONFIG_INVALID_EXCEPTION("FileReader-09", "您的参数配置错误."),
	RUNTIME_EXCEPTION("FileReader-10", "出现运行时异常, 请联系我们"),
	EMPTY_DIR_EXCEPTION("FileReader-11", "您尝试读取的文件目录为空."),;

	private final String code;
	private final String description;

	private FileReaderErrorCode(String code, String description) {
		this.code = code;
		this.description = description;
	}

	@Override
	public String getCode() {
		return this.code;
	}

	@Override
	public String getDescription() {
		return this.description;
	}

	@Override
	public String toString() {
		return String.format("Code:[%s], Description:[%s].", this.code,
				this.description);
	}
}
