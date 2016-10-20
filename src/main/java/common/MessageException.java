package common;

import java.security.InvalidParameterException;

/**
 * 业务异常处理类
 * @author lixiangjing
 */
public class MessageException extends RuntimeException {
	
	private String errCode;

	public String getErrCode() {
		return errCode;
	}

	public MessageException setErrCode(String errCode) {
		this.errCode = errCode;
		return this;
	}
	
	private static final long serialVersionUID = -2282295769920642919L;

	public MessageException(String message) {
		super(message);
	}
	
	
	public static MessageException newMessageException(String message){
		return new MessageException(message);
	}
	
	public static MessageException newMessageException(String message, Object ... args){
		try {
			for(Object arg : args){
				message = message.replace("{}", String.valueOf(arg));
			}
		} catch (Exception e) {
			throw new InvalidParameterException();
		}
		return new MessageException(message);
	}
	

}
